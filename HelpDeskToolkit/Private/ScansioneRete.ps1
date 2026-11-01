# Scansione della rete locale, senza programmi esterni.
# Come fanno gli scanner di rete: ping a tutti gli indirizzi, poi la tabella ARP di Windows
# (trova anche i PC che non rispondono al ping, perché devono comunque rispondere all'ARP),
# poi i nomi via NetBIOS e DNS. Legge e basta: non cambia nulla sui dispositivi.

#region Calcoli sugli indirizzi (funzioni pure)

function ConvertTo-HDNumeroIp {
    param([Parameter(Mandatory)][string]$Ip)
    $b = ([System.Net.IPAddress]::Parse($Ip.Trim())).GetAddressBytes()
    if ($b.Length -ne 4) { throw "Non è un indirizzo IPv4: $Ip" }
    [int64]$b[0] * 16777216 + [int64]$b[1] * 65536 + [int64]$b[2] * 256 + [int64]$b[3]
}

function ConvertFrom-HDNumeroIp {
    param([Parameter(Mandatory)][int64]$Numero)
    '{0}.{1}.{2}.{3}' -f (($Numero -shr 24) -band 255), (($Numero -shr 16) -band 255), (($Numero -shr 8) -band 255), ($Numero -band 255)
}

function Get-HDPrefissoDaMaschera {
    # "255.255.255.0" -> 24
    param([Parameter(Mandatory)][string]$Maschera)
    $n = ConvertTo-HDNumeroIp $Maschera
    $prefisso = 0
    for ($i = 31; $i -ge 0; $i--) {
        if ((($n -shr $i) -band 1) -eq 1) { $prefisso++ } else { break }
    }
    $prefisso
}

function Get-HDInfoRete {
    # "192.168.1.37/24" -> rete 192.168.1.0/24, primo e ultimo host, numero di indirizzi
    param([Parameter(Mandatory)][string]$Rete, [int]$MassimoIndirizzi = 4096)
    if ($Rete -notmatch '^\s*(\d{1,3}(?:\.\d{1,3}){3})\s*/\s*(\d{1,2})\s*$') {
        throw "Rete non valida: '$Rete'. Scrivila così: 192.168.1.0/24"
    }
    $prefisso = [int]$Matches[2]
    if ($prefisso -lt 8 -or $prefisso -gt 30) { throw "Prefisso /$prefisso non valido: usa un valore tra /16 e /30 (di solito /24)." }
    $tutti = [int64]4294967295
    $ip = ConvertTo-HDNumeroIp $Matches[1]
    $maschera = ($tutti -shl (32 - $prefisso)) -band $tutti
    $indirizzoRete = $ip -band $maschera
    $broadcast = $indirizzoRete -bor ((-bnot $maschera) -band $tutti)
    $quanti = $broadcast - $indirizzoRete - 1
    if ($quanti -gt $MassimoIndirizzi) {
        throw "La rete $(ConvertFrom-HDNumeroIp $indirizzoRete)/$prefisso ha $quanti indirizzi: troppi da scansionare (massimo $MassimoIndirizzi). Indica le parti che ti servono, per esempio 10.0.1.0/24, 10.0.2.0/24."
    }
    [pscustomobject]@{
        Rete      = '{0}/{1}' -f (ConvertFrom-HDNumeroIp $indirizzoRete), $prefisso
        Primo     = $indirizzoRete + 1
        Ultimo    = $broadcast - 1
        Indirizzi = $quanti
    }
}

function Get-HDInfoScansione {
    # Cosa scansionare, scritto come nella casella "Rete": una o più reti o intervalli separati da virgole.
    # "192.168.1.0/24, 192.168.2.0/24"   "10.0.0.1-10.0.0.200"   "10.0.0.1-200"   "10.0.0.15"
    param([Parameter(Mandatory)][string]$Testo, [int]$MassimoIndirizzi = 4096)
    $pulito = ($Testo -replace '\s*([/-])\s*', '$1').Trim()
    $pezzi = @($pulito -split '[,;\s]+' | Where-Object { $_ })
    if ($pezzi.Count -eq 0) { throw 'Scrivi la rete da scansionare, per esempio 192.168.1.0/24' }
    $esempi = 'Esempi: 192.168.1.0/24   oppure   192.168.1.0/24, 192.168.2.0/24   oppure   10.0.0.1-10.0.0.200'
    $intervalli = foreach ($pezzo in $pezzi) {
        if (@($pezzo -split '[./-]' | Where-Object { $_ -match '^\d+$' -and [int]$_ -gt 255 }).Count -gt 0) {
            throw "Indirizzo non valido: '$pezzo'. $esempi"
        }
        if ($pezzo -match '^(\d{1,3}(?:\.\d{1,3}){3})-(\d{1,3}(?:\.\d{1,3}){3})$') {
            $da = ConvertTo-HDNumeroIp $Matches[1]
            $a = ConvertTo-HDNumeroIp $Matches[2]
        }
        elseif ($pezzo -match '^((\d{1,3}\.\d{1,3}\.\d{1,3})\.\d{1,3})-(\d{1,3})$') {
            $da = ConvertTo-HDNumeroIp $Matches[1]
            $a = ConvertTo-HDNumeroIp ('{0}.{1}' -f $Matches[2], $Matches[3])
        }
        elseif ($pezzo -match '^\d{1,3}(?:\.\d{1,3}){3}$') {
            $da = ConvertTo-HDNumeroIp $pezzo
            $a = $da
        }
        elseif ($pezzo -match '/') {
            $r = Get-HDInfoRete $pezzo -MassimoIndirizzi $MassimoIndirizzi
            [pscustomobject]@{ Testo = $r.Rete; Primo = $r.Primo; Ultimo = $r.Ultimo }
            continue
        }
        else { throw "Rete non valida: '$pezzo'. $esempi" }
        if ($a -lt $da) { throw "Intervallo non valido: '$pezzo' (il secondo indirizzo deve essere più alto del primo)." }
        [pscustomobject]@{ Testo = $pezzo; Primo = $da; Ultimo = $a }
    }
    $intervalli = @($intervalli)
    $totale = [int64]0
    foreach ($i in $intervalli) { $totale += $i.Ultimo - $i.Primo + 1 }
    if ($totale -gt $MassimoIndirizzi) {
        throw "In tutto sono $totale indirizzi: troppi da scansionare in una volta (massimo $MassimoIndirizzi). Dividi la scansione in più parti."
    }
    [pscustomobject]@{
        Rete       = (@($intervalli | ForEach-Object { $_.Testo }) -join ', ')
        Intervalli = $intervalli
        Primo      = (@($intervalli | ForEach-Object { $_.Primo }) | Measure-Object -Minimum).Minimum
        Ultimo     = (@($intervalli | ForEach-Object { $_.Ultimo }) | Measure-Object -Maximum).Maximum
        Indirizzi  = $totale
    }
}

function Get-HDIntervalliInfo {
    # Get-HDInfoRete ha un solo intervallo (Primo-Ultimo), Get-HDInfoScansione un elenco
    param($Info)
    $elenco = @(@($Info.Intervalli) | Where-Object { $_ })
    if ($elenco.Count -eq 0) { $elenco = @($Info) }
    $elenco
}

function Get-HDIndirizziRete {
    param([Parameter(Mandatory)]$Info)
    $visti = New-Object 'System.Collections.Generic.HashSet[int64]'
    foreach ($i in (Get-HDIntervalliInfo $Info)) {
        for ($n = [int64]$i.Primo; $n -le [int64]$i.Ultimo; $n++) {
            if ($visti.Add($n)) { ConvertFrom-HDNumeroIp $n }
        }
    }
}

function Test-HDIpNellaRete {
    param([string]$Ip, $Info)
    if ($Ip -notmatch '^\d{1,3}(\.\d{1,3}){3}$') { return $false }
    $n = ConvertTo-HDNumeroIp $Ip
    foreach ($i in (Get-HDIntervalliInfo $Info)) {
        if ($n -ge [int64]$i.Primo -and $n -le [int64]$i.Ultimo) { return $true }
    }
    $false
}

function ConvertFrom-HDTabellaArp {
    # Output di "arp -a" (qualsiasi lingua) -> IP e MAC
    param([string[]]$Righe)
    foreach ($r in $Righe) {
        if ($r -match '^\s*(\d{1,3}(?:\.\d{1,3}){3})\s+([0-9a-fA-F]{2}(?:[-:][0-9a-fA-F]{2}){5})\s') {
            $mac = ConvertTo-HDMac $Matches[2]
            if ($mac -and $mac -notmatch '^(01:00:5E|33:33)') { [pscustomobject]@{ IP = $Matches[1]; MAC = $mac } }
        }
    }
}

#endregion

#region NetBIOS (nome del PC e gruppo, come "nbtstat -A")

function New-HDRichiestaNbstat {
    # Richiesta "node status" (RFC 1002) per il nome jolly "*"
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un pacchetto in memoria')]
    param([int]$Id = 0x4844)
    $b = New-Object byte[] 50
    $b[0] = [byte](($Id -shr 8) -band 255)
    $b[1] = [byte]($Id -band 255)
    $b[5] = 1            # una domanda
    $b[12] = 0x20        # lunghezza del nome codificato
    $nome = @(0x2A) + @(0) * 15
    $i = 13
    foreach ($c in $nome) {
        $b[$i] = [byte](0x41 + (($c -shr 4) -band 0x0F))
        $b[$i + 1] = [byte](0x41 + ($c -band 0x0F))
        $i += 2
    }
    $b[47] = 0x21        # tipo NBSTAT
    $b[49] = 0x01        # classe IN
    , $b
}

function ConvertFrom-HDRispostaNbstat {
    # Risposta node status -> nome del computer, gruppo (dominio o gruppo di lavoro), MAC
    param([byte[]]$Byte)
    if (-not $Byte -or $Byte.Length -lt 57) { return $null }
    $pos = 12
    if (($Byte[$pos] -band 0xC0) -eq 0xC0) { $pos += 2 }
    else {
        while ($pos -lt $Byte.Length -and $Byte[$pos] -ne 0) { $pos += 1 + $Byte[$pos] }
        $pos += 1
    }
    $pos += 10           # tipo, classe, TTL, lunghezza dati
    if ($pos -ge $Byte.Length) { return $null }
    $quanti = $Byte[$pos]
    $pos += 1
    $nome = ''
    $gruppo = ''
    for ($k = 0; $k -lt $quanti; $k++) {
        $o = $pos + $k * 18
        if ($o + 18 -gt $Byte.Length) { break }
        $testo = [System.Text.Encoding]::ASCII.GetString($Byte, $o, 15).Trim([char[]]@([char]0, [char]32))
        $suffisso = $Byte[$o + 15]
        $eGruppo = ($Byte[$o + 16] -band 0x80) -ne 0
        if ($suffisso -eq 0 -and -not $eGruppo -and -not $nome) { $nome = $testo }
        if ($suffisso -eq 0 -and $eGruppo -and -not $gruppo) { $gruppo = $testo }
    }
    $mac = ''
    $o = $pos + $quanti * 18
    if ($o + 6 -le $Byte.Length) { $mac = ConvertTo-HDMac ((0..5 | ForEach-Object { '{0:X2}' -f $Byte[$o + $_] }) -join '') }
    if (-not $nome -and -not $gruppo) { return $null }
    [pscustomobject]@{ Nome = $nome; Gruppo = $gruppo; MAC = $mac }
}

#endregion

#region Operazioni di rete (solo Windows)

function Get-HDReteLocale {
    # Rete della scheda principale: IP e rete da scansionare (tutta fino a /20, altrimenti la /24 del PC)
    $config = @(Get-HDCim -Classe Win32_NetworkAdapterConfiguration -Filtro 'IPEnabled=True' | Select-Object Description, IPAddress, IPSubnet, MACAddress, DefaultIPGateway, DHCPEnabled, IPConnectionMetric)
    $principale = Select-HDSchedaPrincipale $config
    if (-not $principale) { throw 'Questo PC non è collegato a una rete (nessun indirizzo IPv4).' }
    $indice = [array]::IndexOf(@($principale.Config.IPAddress), $principale.IP)
    $maschera = @($principale.Config.IPSubnet)[$indice]
    $prefisso = 24
    if ($maschera -match '^\d{1,3}(\.\d{1,3}){3}$') { $prefisso = Get-HDPrefissoDaMaschera $maschera }
    $nota = ''
    if ($prefisso -lt 20) {
        $nota = "La rete di questo PC è molto grande (/$prefisso): propongo solo la sua /24. Per scansionarne di più aggiungi altre reti separate da virgole (fino a 4096 indirizzi in tutto)."
        $prefisso = 24
    }
    [pscustomobject]@{
        IP       = $principale.IP
        MAC      = (ConvertTo-HDMac $principale.Config.MACAddress)
        Rete     = (Get-HDInfoRete ('{0}/{1}' -f $principale.IP, $prefisso)).Rete
        Nota     = $nota
        Hostname = (Get-HDNomePC).ToUpperInvariant()
    }
}

function Invoke-HDPingParallelo {
    # Ping a blocchi di indirizzi in parallelo; restituisce l'elenco di quelli che rispondono
    param([Parameter(Mandatory)][string[]]$Indirizzi, [int]$TimeoutMs = 800, [int]$Blocco = 128, [scriptblock]$Avanzamento)
    $vivi = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Indirizzi.Count; $i += $Blocco) {
        $parte = @($Indirizzi[$i..([math]::Min($i + $Blocco, $Indirizzi.Count) - 1)])
        $prove = foreach ($ip in $parte) {
            $p = New-Object System.Net.NetworkInformation.Ping
            [pscustomobject]@{ IP = $ip; Ping = $p; Attesa = $p.SendPingAsync($ip, $TimeoutMs) }
        }
        try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($prove | ForEach-Object { $_.Attesa }), $TimeoutMs + 3000) }
        catch { Write-Verbose "Alcuni ping non sono partiti: $($_.Exception.Message)" }
        foreach ($p in $prove) {
            if ($p.Attesa.Status -eq 'RanToCompletion' -and $p.Attesa.Result.Status -eq 'Success') { $vivi.Add($p.IP) }
            $p.Ping.Dispose()
        }
        if ($Avanzamento) { & $Avanzamento ('Ping: {0} di {1} indirizzi' -f [math]::Min($i + $Blocco, $Indirizzi.Count), $Indirizzi.Count) }
    }
    $vivi.ToArray()
}

function Get-HDViciniArp {
    # Tabella ARP di Windows (dopo il ping contiene anche chi blocca il ping)
    param([Parameter(Mandatory)]$Info)
    $voci = @()
    if (Get-Command -Name Get-NetNeighbor -ErrorAction SilentlyContinue) {
        try {
            $voci = @(Get-NetNeighbor -AddressFamily IPv4 -ErrorAction Stop |
                    Where-Object { [string]$_.State -notin 'Unreachable', 'Incomplete', 'Permanent' } |
                    ForEach-Object { [pscustomobject]@{ IP = [string]$_.IPAddress; MAC = (ConvertTo-HDMac $_.LinkLayerAddress) } } |
                    Where-Object { $_.MAC })
        }
        catch { Write-Verbose "Get-NetNeighbor non disponibile: $($_.Exception.Message)" }
    }
    if ($voci.Count -eq 0) { $voci = @(ConvertFrom-HDTabellaArp -Righe @(& arp.exe -a)) }
    @($voci | Where-Object { Test-HDIpNellaRete -Ip $_.IP -Info $Info })
}

function Resolve-HDNetbiosRemoto {
    param([Parameter(Mandatory)][string[]]$Indirizzi, [int]$AttesaMs = 1500)
    $risposte = @{}
    $udp = New-Object System.Net.Sockets.UdpClient(0)
    try {
        $udp.Client.ReceiveTimeout = 250
        $udp.Client.ReceiveBufferSize = 1048576   # reti grandi: tante risposte tutte insieme
        $richiesta = New-HDRichiestaNbstat
        foreach ($ip in $Indirizzi) {
            try { [void]$udp.Send($richiesta, $richiesta.Length, $ip, 137) } catch { Write-Verbose "NetBIOS non inviato a ${ip}: $($_.Exception.Message)" }
        }
        $fine = (Get-Date).AddMilliseconds($AttesaMs)
        $da = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        while ((Get-Date) -lt $fine -and $risposte.Count -lt $Indirizzi.Count) {
            try {
                $dati = $udp.Receive([ref]$da)
                $info = ConvertFrom-HDRispostaNbstat -Byte $dati
                if ($info) { $risposte[$da.Address.ToString()] = $info }
            }
            catch { Write-Verbose 'Attesa risposte NetBIOS...' }
        }
    }
    finally { $udp.Close() }
    $risposte
}

function Resolve-HDNomeDnsInverso {
    param([Parameter(Mandatory)][string[]]$Indirizzi, [int]$TimeoutMs = 0)
    $nomi = @{}
    if ($Indirizzi.Count -eq 0) { return $nomi }
    if ($TimeoutMs -le 0) { $TimeoutMs = [math]::Min(15000, 2500 + 10 * $Indirizzi.Count) }
    $richieste = foreach ($ip in $Indirizzi) {
        [pscustomobject]@{ IP = $ip; Attesa = [System.Net.Dns]::GetHostEntryAsync($ip) }
    }
    try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($richieste | ForEach-Object { $_.Attesa }), $TimeoutMs) }
    catch { Write-Verbose 'Alcuni nomi DNS non risolti' }
    foreach ($r in $richieste) {
        if ($r.Attesa.Status -eq 'RanToCompletion' -and $r.Attesa.Result.HostName -and $r.Attesa.Result.HostName -ne $r.IP) { $nomi[$r.IP] = $r.Attesa.Result.HostName }
    }
    $nomi
}

#endregion

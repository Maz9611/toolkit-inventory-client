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

#region Produttore dal MAC (elenco IEEE) e nomi mDNS

function Get-HDProduttoreMac {
    # "EC:71:DB:AE:F6:2F" -> "Reolink Innovation Limited"; i MAC casuali di telefoni e tablet non hanno produttore
    param([string]$Mac)
    $m = ConvertTo-HDMac $Mac
    if (-not $m) { return '' }
    $esadecimale = $m -replace ':', ''
    if (([Convert]::ToInt32($esadecimale.Substring(0, 2), 16) -band 2) -ne 0) { return 'MAC casuale (telefono, tablet o PC con privacy attiva)' }
    if ($null -eq $script:HDProduttoriMac) { $script:HDProduttoriMac = Read-HDElencoProduttoriMac }
    # Prima i blocchi piccoli (MA-S: 9 cifre, MA-M: 7), poi l'OUI (6). Un OUI dell'IEEE Registration
    # Authority è diviso tra tanti produttori: se il blocco non è nell'elenco, il produttore non si sa
    $nome = $null
    foreach ($cifre in 9, 7, 6) {
        if ($script:HDProduttoriMac.TryGetValue($esadecimale.Substring(0, $cifre), [ref]$nome)) {
            if ($nome -eq 'IEEE Registration Authority') { return '' }
            return $nome
        }
    }
    ''
}

function Read-HDElencoProduttoriMac {
    # Elenco "prefisso<TAB>produttore" compresso (prefisso di 6, 7 o 9 cifre esadecimali): dal file accanto
    # al modulo oppure (nell'exe) dal testo passato all'avvio
    $elenco = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    $testo = $script:HDTestoProduttoriMac
    if (-not $testo) {
        $file = Join-Path (Join-Path $script:HDModuloRadice 'Dati') 'produttori-mac.txt.gz'
        if (-not (Test-Path -LiteralPath $file)) { return $elenco }
        $flusso = [System.IO.File]::OpenRead($file)
        try {
            $gz = New-Object System.IO.Compression.GZipStream($flusso, [System.IO.Compression.CompressionMode]::Decompress)
            $testo = (New-Object System.IO.StreamReader($gz, [System.Text.Encoding]::UTF8)).ReadToEnd()
        }
        finally { $flusso.Dispose() }
    }
    foreach ($riga in $testo.Split([char]10)) {
        $i = $riga.IndexOf([char]9)
        if ($i -eq 6 -or $i -eq 7 -or $i -eq 9) { $elenco[$riga.Substring(0, $i)] = $riga.Substring($i + 1).Trim() }
    }
    $elenco
}

# Servizi chiesti al gruppo multicast: chi li offre risponde con il suo nome (record A nome.local -> IP)
$script:HDServiziMdns = @(
    '_companion-link._tcp.local', '_rdlink._tcp.local', '_apple-mobdev2._tcp.local'   # Mac, iPhone, iPad
    '_airplay._tcp.local', '_raop._tcp.local', '_sleep-proxy._udp.local'              # Apple TV, HomePod, AirPlay
    '_smb._tcp.local', '_afpovertcp._tcp.local', '_workstation._tcp.local', '_ssh._tcp.local'
    '_ipp._tcp.local', '_ipps._tcp.local', '_printer._tcp.local', '_pdl-datastream._tcp.local', '_uscan._tcp.local'
    '_googlecast._tcp.local', '_spotify-connect._tcp.local', '_hap._tcp.local', '_http._tcp.local'
)

function New-HDDomandaMulticast {
    # Domanda DNS con uno o più nomi (PTR se non indicato). -Unicast chiede la risposta diretta (bit QU):
    # con la porta di partenza diversa da 5353 i dispositivi rispondono comunque solo a chi ha chiesto
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un pacchetto in memoria')]
    param([Parameter(Mandatory)][string[]]$Nomi, [int]$Tipo = 12, [int]$Id = 0x4844, [switch]$Unicast)
    $classe = 1
    if ($Unicast) { $classe = 0x8001 }
    $byte = New-Object System.Collections.Generic.List[byte]
    foreach ($b in @((($Id -shr 8) -band 255), ($Id -band 255), 0, 0, (($Nomi.Count -shr 8) -band 255), ($Nomi.Count -band 255), 0, 0, 0, 0, 0, 0)) { $byte.Add([byte]$b) }
    foreach ($nome in $Nomi) {
        foreach ($etichetta in $nome.Trim('.').Split('.')) {
            $testo = [System.Text.Encoding]::UTF8.GetBytes($etichetta)
            $byte.Add([byte]$testo.Length)
            $byte.AddRange($testo)
        }
        foreach ($b in @(0, (($Tipo -shr 8) -band 255), ($Tipo -band 255), (($classe -shr 8) -band 255), ($classe -band 255))) { $byte.Add([byte]$b) }
    }
    , $byte.ToArray()
}

function ConvertTo-HDNomeInverso {
    # 192.168.3.184 -> 184.3.168.192.in-addr.arpa
    param([string]$Ip)
    $p = $Ip.Split('.')
    '{0}.{1}.{2}.{3}.in-addr.arpa' -f $p[3], $p[2], $p[1], $p[0]
}

function New-HDPacchettoMdnsInverso {
    # Domanda "chi è x.x.x.x?" (PTR di in-addr.arpa) per uno o più indirizzi
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un pacchetto in memoria')]
    param([Parameter(Mandatory)][string[]]$Ip, [int]$Id = 0x4844, [switch]$Unicast)
    New-HDDomandaMulticast -Nomi @($Ip | ForEach-Object { ConvertTo-HDNomeInverso $_ }) -Id $Id -Unicast:$Unicast
}

function Get-HDDomandeMulticast {
    # Pacchetti per il gruppo multicast: i servizi più comuni, poi i nomi inversi degli indirizzi a gruppi
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo pacchetti in memoria')]
    param([string[]]$Indirizzi, [int]$PerPacchetto = 30)
    , (New-HDDomandaMulticast -Nomi $script:HDServiziMdns -Id 0x4845 -Unicast)
    $elenco = @($Indirizzi | Where-Object { $_ })
    for ($i = 0; $i -lt $elenco.Count; $i += $PerPacchetto) {
        $ultimo = [math]::Min($i + $PerPacchetto, $elenco.Count) - 1
        , (New-HDPacchettoMdnsInverso -Ip $elenco[$i..$ultimo] -Id 0x4846 -Unicast)
    }
}

function ConvertFrom-HDNomeDominio {
    # Nome DNS a partire da una posizione del pacchetto (anche compresso): Nome e posizione dopo il nome
    param([byte[]]$Byte, [int]$Posizione)
    $parti = New-Object System.Collections.Generic.List[string]
    $pos = $Posizione
    $fine = -1
    $salti = 0
    while ($pos -lt $Byte.Length) {
        $lunghezza = [int]$Byte[$pos]
        if ($lunghezza -eq 0) { $pos++; break }
        if (($lunghezza -band 0xC0) -eq 0xC0) {
            if ($pos + 1 -ge $Byte.Length) { break }
            if ($fine -lt 0) { $fine = $pos + 2 }
            $pos = (($lunghezza -band 0x3F) * 256) + [int]$Byte[$pos + 1]
            $salti++
            if ($salti -gt 20) { break }
            continue
        }
        if ($pos + 1 + $lunghezza -gt $Byte.Length) { break }
        $parti.Add([System.Text.Encoding]::UTF8.GetString($Byte, $pos + 1, $lunghezza))
        $pos += 1 + $lunghezza
    }
    if ($fine -lt 0) { $fine = $pos }
    [pscustomobject]@{ Nome = ($parti -join '.'); Fine = $fine }
}

function ConvertFrom-HDRecordMulticast {
    # Record di una risposta mDNS (risposte, autorità e sezione aggiuntiva) di tipo A, PTR, CNAME e SRV:
    # Nome, Tipo e Valore (l'IP per A, il nome indicato per gli altri)
    param([byte[]]$Byte)
    if (-not $Byte -or $Byte.Length -lt 12) { return }
    $domande = [int]$Byte[4] * 256 + [int]$Byte[5]
    $record = 0
    foreach ($i in 6, 8, 10) { $record += [int]$Byte[$i] * 256 + [int]$Byte[$i + 1] }
    $pos = 12
    for ($i = 0; $i -lt $domande -and $pos -lt $Byte.Length; $i++) { $pos = (ConvertFrom-HDNomeDominio $Byte $pos).Fine + 4 }
    for ($i = 0; $i -lt $record -and $pos + 10 -le $Byte.Length; $i++) {
        $letto = ConvertFrom-HDNomeDominio $Byte $pos
        $pos = $letto.Fine
        if ($pos + 10 -gt $Byte.Length) { break }
        $tipo = [int]$Byte[$pos] * 256 + [int]$Byte[$pos + 1]
        $lunghezza = [int]$Byte[$pos + 8] * 256 + [int]$Byte[$pos + 9]
        $inizio = $pos + 10
        if ($inizio + $lunghezza -gt $Byte.Length) { break }
        $valore = ''
        if ($tipo -eq 1 -and $lunghezza -eq 4) { $valore = '{0}.{1}.{2}.{3}' -f $Byte[$inizio], $Byte[$inizio + 1], $Byte[$inizio + 2], $Byte[$inizio + 3] }
        elseif ($tipo -eq 12 -or $tipo -eq 5) { $valore = (ConvertFrom-HDNomeDominio $Byte $inizio).Nome }
        elseif ($tipo -eq 33 -and $lunghezza -gt 6) { $valore = (ConvertFrom-HDNomeDominio $Byte ($inizio + 6)).Nome }
        if ($valore) { [pscustomobject]@{ Nome = $letto.Nome; Tipo = $tipo; Valore = $valore } }
        $pos = $inizio + $lunghezza
    }
}

function Get-HDNomiDaMulticast {
    # Nomi dei dispositivi in una risposta mDNS (IP -> nome senza ".local"): dai record A (nome.local -> IP)
    # e dai PTR inversi (x.x.x.x.in-addr.arpa -> nome.local). Se la risposta dice solo il nome host del
    # servizio (SRV) senza il suo indirizzo, il nome vale per chi ha risposto ($Da)
    param([byte[]]$Byte, [string]$Da)
    $nomi = @{}
    $conIndirizzo = @{}
    $record = @(ConvertFrom-HDRecordMulticast -Byte $Byte)
    foreach ($r in $record) {
        $ip = ''
        $nome = ''
        if ($r.Tipo -eq 1) { $ip = $r.Valore; $nome = $r.Nome }
        elseif ($r.Tipo -eq 12 -and $r.Nome -match '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.in-addr\.arpa\.?$') {
            $ip = '{0}.{1}.{2}.{3}' -f $Matches[4], $Matches[3], $Matches[2], $Matches[1]
            $nome = $r.Valore
        }
        $nome = ($nome -replace '(?i)\.local\.?$', '').Trim('.')
        if (-not ($ip -and $nome)) { continue }
        $conIndirizzo[$nome.ToLowerInvariant()] = $true
        if (-not $nomi.ContainsKey($ip)) { $nomi[$ip] = $nome }
    }
    if ($Da -and -not $nomi.ContainsKey($Da)) {
        foreach ($r in @($record | Where-Object { $_.Tipo -eq 33 })) {
            $nome = ($r.Valore -replace '(?i)\.local\.?$', '').Trim('.')
            if ($nome -and -not $conIndirizzo.ContainsKey($nome.ToLowerInvariant())) { $nomi[$Da] = $nome; break }
        }
    }
    $nomi
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
    param([Parameter(Mandatory)][string[]]$Indirizzi, [int]$TimeoutMs = 800, [int]$Blocco = 256, [scriptblock]$Avanzamento)
    $vivi = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Indirizzi.Count; $i += $Blocco) {
        $parte = @($Indirizzi[$i..([math]::Min($i + $Blocco, $Indirizzi.Count) - 1)])
        $prove = foreach ($ip in $parte) {
            $p = New-Object System.Net.NetworkInformation.Ping
            [pscustomobject]@{ IP = $ip; Ping = $p; Attesa = $p.SendPingAsync($ip, $TimeoutMs) }
        }
        # Windows può metterci qualche secondo a dare per irraggiungibile un indirizzo senza risposta ARP:
        # non si aspetta oltre il timeout (chi risponde lo fa entro quel tempo)
        try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]@($prove | ForEach-Object { $_.Attesa }), $TimeoutMs + 700) }
        catch { Write-Verbose "Alcuni ping non sono partiti: $($_.Exception.Message)" }
        foreach ($p in $prove) {
            if ($p.Attesa.Status -eq 'RanToCompletion' -and $p.Attesa.Result.Status -eq 'Success') { $vivi.Add($p.IP) }
            if ($p.Attesa.IsCompleted) { $p.Ping.Dispose() }
        }
        if ($Avanzamento) { & $Avanzamento ('Ping: {0} di {1} indirizzi' -f [math]::Min($i + $Blocco, $Indirizzi.Count), $Indirizzi.Count) }
    }
    $vivi.ToArray()
}

function Invoke-HDSondaTcp {
    # Chi non risponde al ping ma esiste: una porta TCP comune che accetta la connessione (Aperta)
    # o la rifiuta. Mac, iPhone, PC con il firewall che blocca il ping. Si apre e si chiude subito
    # la connessione: nessun dato inviato. Restituisce IP e Aperta ($true se almeno una porta è aperta);
    # chi usa il risultato decide se un rifiuto basta (sulla rete del PC sì, oltre un router no).
    param(
        [Parameter(Mandatory)][string[]]$Indirizzi,
        [int[]]$Porte = @(445, 80, 443, 22, 7000, 62078),
        [int]$TimeoutMs = 1500,
        [int]$Blocco = 512,
        [scriptblock]$Avanzamento
    )
    $vivi = @{}
    $coppie = @(foreach ($ip in $Indirizzi) { foreach ($porta in $Porte) { [pscustomobject]@{ IP = $ip; Porta = $porta } } })
    for ($i = 0; $i -lt $coppie.Count; $i += $Blocco) {
        $parte = @($coppie[$i..([math]::Min($i + $Blocco, $coppie.Count) - 1)])
        $prove = foreach ($c in $parte) {
            $client = New-Object System.Net.Sockets.TcpClient
            $attesa = $null
            try { $attesa = $client.ConnectAsync($c.IP, $c.Porta) } catch { $attesa = $null }
            [pscustomobject]@{ IP = $c.IP; Client = $client; Attesa = $attesa }
        }
        $compiti = @($prove | Where-Object { $_.Attesa } | ForEach-Object { $_.Attesa })
        if ($compiti.Count -gt 0) {
            try { [void][System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]$compiti, $TimeoutMs) }
            catch { Write-Verbose 'Alcune porte hanno rifiutato o non hanno risposto' }
        }
        foreach ($p in $prove) {
            if ($p.Attesa -and $p.Attesa.IsCompleted) {
                if ($p.Attesa.Status -eq 'RanToCompletion') { $vivi[$p.IP] = $true }
                elseif ($p.Attesa.Exception) {
                    $errore = $p.Attesa.Exception.InnerException
                    if ($errore -is [System.Net.Sockets.SocketException] -and [string]$errore.SocketErrorCode -eq 'ConnectionRefused' -and -not $vivi.ContainsKey($p.IP)) { $vivi[$p.IP] = $false }
                }
            }
            try { $p.Client.Close() } catch { Write-Verbose 'Connessione già chiusa' }
        }
        if ($Avanzamento) { & $Avanzamento ('Controllo delle porte: {0} di {1}' -f [math]::Min($i + $Blocco, $coppie.Count), $coppie.Count) }
    }
    foreach ($ip in $vivi.Keys) { [pscustomobject]@{ IP = $ip; Aperta = [bool]$vivi[$ip] } }
}

function Select-HDSondeAffidabili {
    # Dal controllo delle porte tiene solo i risultati credibili:
    # - sulla rete del PC basta un rifiuto (lì può rispondere solo il dispositivo stesso);
    # - oltre un router serve una porta aperta (un firewall può rifiutare al posto di tutti);
    # - se risultano "vivi" quasi tutti gli indirizzi provati, qualcosa risponde per tutti: si scarta.
    [CmdletBinding()]
    param([object[]]$Sonde, $InfoLocale, [int]$Provati)
    $tenute = @(@($Sonde | Where-Object { $_ }) | Where-Object { $_.Aperta -or ($InfoLocale -and (Test-HDIpNellaRete -Ip $_.IP -Info $InfoLocale)) })
    $sospetto = { param($n) $n -gt 20 -and $n -gt ($Provati * 0.5) }
    if (& $sospetto $tenute.Count) {
        $tenute = @($tenute | Where-Object { $_.Aperta })
        if (& $sospetto $tenute.Count) { $tenute = @() }
        Write-Warning 'Troppi indirizzi rispondono sulle porte TCP (probabilmente un firewall risponde al posto di tutti): quel controllo non viene usato.'
    }
    @($tenute | ForEach-Object { $_.IP })
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

function Test-HDNomeValido {
    # Scarta i nomi segnaposto che alcuni dispositivi annunciano ("(none)", "localhost"...)
    param([string]$Nome)
    $n = ([string]$Nome).Trim()
    ($n -and $n -notmatch '(?i)^(\(none\).*|none|localhost|unknown|null|default|-+)$')
}

function New-HDClientUdp {
    # Client UDP su una porta qualsiasi, per le domande NetBIOS e mDNS a tanti indirizzi. Su Windows il
    # "porta chiusa" (ICMP) di un dispositivo fa fallire la lettura successiva (WSAECONNRESET): con
    # centinaia di domande quei falsi errori consumerebbero il tempo di attesa delle risposte vere
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param()
    $udp = New-Object System.Net.Sockets.UdpClient(0)
    $udp.Client.ReceiveTimeout = 250
    $udp.Client.ReceiveBufferSize = 1048576   # reti grandi: tante risposte tutte insieme
    if (Test-HDWindows) {
        # SIO_UDP_CONNRESET (0x9800000C) = falso
        try { [void]$udp.Client.IOControl(-1744830452, [byte[]]@(0, 0, 0, 0), $null) }
        catch { Write-Verbose "SIO_UDP_CONNRESET non impostato: $($_.Exception.Message)" }
    }
    $udp
}

function Resolve-HDNetbiosRemoto {
    param([Parameter(Mandatory)][string[]]$Indirizzi, [int]$AttesaMs = 1500)
    $risposte = @{}
    $udp = New-HDClientUdp
    try {
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

function Resolve-HDNomeMulticast {
    # Chiede il nome via mDNS (stampanti, telefoni, router, dispositivi Apple/Android, IoT): domanda diretta
    # a ogni indirizzo e, per gli indirizzi della rete del PC (-Locali), domande al gruppo multicast
    # 224.0.0.251 dalla scheda di -IpLocale: Mac, iPhone e iPad spesso rispondono solo a queste
    param([Parameter(Mandatory)][string[]]$Indirizzi, [string]$IpLocale, [string[]]$Locali = @(), [int]$AttesaMs = 1500)
    $nomi = @{}
    $cercati = @{}
    foreach ($ip in $Indirizzi) { $cercati[$ip] = $true }
    $udp = New-HDClientUdp
    try {
        foreach ($ip in $Indirizzi) {
            $richiesta = New-HDPacchettoMdnsInverso -Ip $ip
            try { [void]$udp.Send($richiesta, $richiesta.Length, $ip, 5353) } catch { Write-Verbose "mDNS non inviato a ${ip}: $($_.Exception.Message)" }
        }
        $perMulticast = @($Locali | Where-Object { $_ -and $cercati.ContainsKey($_) })
        if ($IpLocale -and $perMulticast.Count -gt 0) {
            try {
                $udp.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::IP, [System.Net.Sockets.SocketOptionName]::MulticastInterface, ([System.Net.IPAddress]::Parse($IpLocale)).GetAddressBytes())
                $udp.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::IP, [System.Net.Sockets.SocketOptionName]::MulticastTimeToLive, 255)
                $gruppo = New-Object System.Net.IPEndPoint(([System.Net.IPAddress]::Parse('224.0.0.251')), 5353)
                # Oltre 1024 indirizzi solo le domande sui servizi: troppe domande inverse intaserebbero il Wi-Fi
                if ($perMulticast.Count -gt 1024) { $perMulticast = @() }
                foreach ($pacchetto in @(Get-HDDomandeMulticast -Indirizzi $perMulticast)) { [void]$udp.Send($pacchetto, $pacchetto.Length, $gruppo) }
            }
            catch { Write-Verbose "mDNS multicast non inviato: $($_.Exception.Message)" }
        }
        $fine = (Get-Date).AddMilliseconds($AttesaMs)
        $da = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        while ((Get-Date) -lt $fine -and $nomi.Count -lt $cercati.Count) {
            try {
                $dati = $udp.Receive([ref]$da)
                $risposta = Get-HDNomiDaMulticast -Byte $dati -Da $da.Address.ToString()
                foreach ($ip in @($risposta.Keys)) {
                    if ($cercati.ContainsKey($ip) -and -not $nomi.ContainsKey($ip)) { $nomi[$ip] = $risposta[$ip] }
                }
            }
            catch { Write-Verbose 'Attesa risposte mDNS...' }
        }
    }
    finally { $udp.Close() }
    $nomi
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

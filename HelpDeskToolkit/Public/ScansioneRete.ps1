# Scansione della rete integrata (al posto di Advanced IP Scanner).

function Invoke-ScansioneRete {
    <#
    .SYNOPSIS
    Scansiona la rete locale e, con -NelCensimento, mette i dispositivi trovati nell'Excel.

    .DESCRIPTION
    Senza programmi esterni: ping a tutti gli indirizzi della rete (due giri), tabella ARP di
    Windows, per chi ancora manca alcune porte TCP comuni (445, 80, 443, 22, 7000, 62078: la
    connessione si apre e si chiude subito), domande NetBIOS e mDNS a tutti gli indirizzi, nomi
    anche dal DNS. Trova così anche Mac, iPhone e PC con il firewall che blocca il ping. Ci
    vogliono di solito 15-20 secondi per una rete /24. Ogni scansione viene salvata anche in
    Censimento\scansioni come CSV.
    Con -NelCensimento i dispositivi entrano nell'Excel con le stesse regole dell'import
    di Advanced IP Scanner: IP sempre aggiornato, il resto solo nelle celle vuote, nessun
    doppione, le colonne che compili tu non vengono toccate.

    Nota: alcuni antivirus/EDR (Cynet compreso) possono segnalare una scansione di rete:
    avvisa prima chi gestisce la sicurezza del cliente.

    .PARAMETER Rete
    Cosa scansionare: una rete (192.168.1.0/24), più reti separate da virgole
    (192.168.1.0/24, 192.168.2.0/24) o intervalli (10.0.0.1-10.0.0.200). Al massimo 4096
    indirizzi in tutto. Se manca, la rete del PC.

    .EXAMPLE
    Invoke-ScansioneRete

    .EXAMPLE
    Invoke-ScansioneRete -NelCensimento -Cliente 'Studio Rossi'

    .EXAMPLE
    Invoke-ScansioneRete -Rete '192.168.10.0/24, 192.168.20.0/24' -NelCensimento

    .NOTES
    Le reti dietro un router (altre VLAN) si scansionano, ma lì la tabella ARP non serve:
    si trovano solo i dispositivi che rispondono al ping e il MAC lo aggiunge poi il
    censimento sul PC.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0)][string]$Rete,
        [switch]$NelCensimento,
        [string]$Censimento,
        [string]$Cliente,
        [int]$TimeoutMs = 800,
        [scriptblock]$Avanzamento,
        [switch]$NonChiedere
    )
    if (-not (Test-HDWindows)) { throw 'La scansione di rete funziona solo su Windows.' }
    $segnala = {
        param($Testo)
        Write-Verbose $Testo
        if ($Avanzamento) { & $Avanzamento $Testo }
    }
    $locale = $null
    try { $locale = Get-HDReteLocale } catch { if (-not $Rete) { throw } }
    if (-not $Rete) {
        $Rete = $locale.Rete
        if ($locale.Nota) { Write-Warning $locale.Nota }
    }
    $info = Get-HDInfoScansione $Rete
    $inizio = Get-Date
    & $segnala ('Scansione di {0} ({1} indirizzi)...' -f $info.Rete, $info.Indirizzi)

    $indirizzi = @(Get-HDIndirizziRete $info)
    $rispondono = @(Invoke-HDPingParallelo -Indirizzi $indirizzi -TimeoutMs $TimeoutMs -Avanzamento $Avanzamento)
    & $segnala 'Lettura della tabella ARP...'
    Start-Sleep -Milliseconds 800
    $arp = @{}
    foreach ($v in @(Get-HDViciniArp -Info $info)) { $arp[$v.IP] = $v.MAC }
    # Secondo giro, più lento, per chi non ha risposto (telefoni e dispositivi a risparmio energetico)
    $mancanti = @($indirizzi | Where-Object { $rispondono -notcontains $_ -and -not $arp.ContainsKey($_) })
    if ($mancanti.Count -gt 0) {
        & $segnala ('Secondo giro per {0} indirizzi che non hanno risposto...' -f $mancanti.Count)
        $rispondono = @($rispondono) + @(Invoke-HDPingParallelo -Indirizzi $mancanti -TimeoutMs ([math]::Max($TimeoutMs * 2, 1500)) -Avanzamento $Avanzamento)
        Start-Sleep -Milliseconds 800
        foreach ($v in @(Get-HDViciniArp -Info $info)) { $arp[$v.IP] = $v.MAC }
    }

    # Terzo controllo per chi ancora manca: porte TCP comuni (Mac, iPhone, PC con firewall)
    $infoLocale = $null
    if ($locale -and $locale.Rete) { try { $infoLocale = Get-HDInfoRete $locale.Rete } catch { $infoLocale = $null } }
    $tcp = @()
    $mancanti = @($indirizzi | Where-Object { $rispondono -notcontains $_ -and -not $arp.ContainsKey($_) })
    if ($mancanti.Count -gt 0 -and $mancanti.Count -le 1024) {
        & $segnala ('Controllo di alcune porte comuni su {0} indirizzi...' -f $mancanti.Count)
        $tcp = @(Select-HDSondeAffidabili -Sonde @(Invoke-HDSondaTcp -Indirizzi $mancanti -Avanzamento $Avanzamento) -InfoLocale $infoLocale -Provati $mancanti.Count)
    }

    # Nomi chiesti a tutti gli indirizzi: chi risponde a NetBIOS o mDNS esiste, anche se ignora il ping
    & $segnala 'Ricerca dei nomi (NetBIOS e mDNS)...'
    $netbios = Resolve-HDNetbiosRemoto -Indirizzi $indirizzi
    # mDNS anche al gruppo multicast, ma solo per la rete del PC: oltre un router non arriva
    $ipLocale = ''
    $locali = @()
    if ($infoLocale) {
        $ipLocale = $locale.IP
        $locali = @($indirizzi | Where-Object { Test-HDIpNellaRete -Ip $_ -Info $infoLocale })
    }
    $mdns = Resolve-HDNomeMulticast -Indirizzi $indirizzi -IpLocale $ipLocale -Locali $locali

    # Ultima lettura della tabella ARP: il MAC anche di chi si è fatto vivo con gli ultimi controlli
    foreach ($v in @(Get-HDViciniArp -Info $info)) { if (-not $arp.ContainsKey($v.IP)) { $arp[$v.IP] = $v.MAC } }

    $trovati = New-Object System.Collections.Generic.List[string]
    foreach ($ip in @($rispondono) + @($arp.Keys) + @($tcp) + @($netbios.Keys) + @($mdns.Keys)) {
        if ($ip -and (Test-HDIpNellaRete -Ip $ip -Info $info) -and -not $trovati.Contains($ip)) { $trovati.Add($ip) }
    }
    if ($locale -and (Test-HDIpNellaRete -Ip $locale.IP -Info $info) -and -not $trovati.Contains($locale.IP)) { $trovati.Add($locale.IP) }

    $dns = @{}
    $senzaNome = @($trovati | Where-Object { -not ($netbios.ContainsKey($_) -and (Test-HDNomeValido $netbios[$_].Nome)) })
    if ($senzaNome.Count -gt 0) {
        & $segnala ('Ricerca dei nomi DNS di {0} dispositivi...' -f $senzaNome.Count)
        $dns = Resolve-HDNomeDnsInverso -Indirizzi $senzaNome
    }

    $dispositivi = foreach ($ip in ($trovati | Sort-Object { ConvertTo-HDNumeroIp $_ })) {
        $nome = ''
        $gruppo = ''
        $mac = $arp[$ip]
        $fonte = @()
        if ($rispondono -contains $ip) { $fonte += 'ping' }
        if ($arp.ContainsKey($ip)) { $fonte += 'ARP' }
        if ($tcp -contains $ip) { $fonte += 'TCP' }
        if ($netbios.ContainsKey($ip) -and (Test-HDNomeValido $netbios[$ip].Nome)) {
            $nome = $netbios[$ip].Nome
            $gruppo = $netbios[$ip].Gruppo
            if (-not $mac) { $mac = $netbios[$ip].MAC }
            $fonte += 'NetBIOS'
        }
        elseif ($dns.ContainsKey($ip) -and (Test-HDNomeValido $dns[$ip])) { $nome = $dns[$ip]; $fonte += 'DNS' }
        elseif ($mdns.ContainsKey($ip) -and (Test-HDNomeValido $mdns[$ip])) { $nome = $mdns[$ip]; $fonte += 'mDNS' }
        elseif ($netbios.ContainsKey($ip) -or $mdns.ContainsKey($ip)) { $fonte += 'NetBIOS/mDNS' }
        if ($locale -and $ip -eq $locale.IP) {
            if (-not $nome) { $nome = $locale.Hostname }
            if (-not $mac) { $mac = $locale.MAC }
            $fonte += 'questo PC'
        }
        [pscustomobject]@{
            PSTypeName = 'HelpDesk.DispositivoRete'
            IP         = $ip
            Hostname   = (ConvertTo-HDNomeHost $nome)
            Nome       = $nome
            MAC        = (ConvertTo-HDMac $mac)
            Gruppo     = $gruppo
            Produttore = (Get-HDProduttoreMac $mac)
            Utente     = ''
            Stato      = 'Attivo'
            Commenti   = ''
            Fonte      = ($fonte -join ', ')
        }
    }
    $dispositivi = @($dispositivi)
    $secondi = [math]::Round(((Get-Date) - $inizio).TotalSeconds)
    & $segnala ('Trovati {0} dispositivi in {1} secondi.' -f $dispositivi.Count, $secondi)

    $copia = $null
    if ($dispositivi.Count -gt 0 -and $PSCmdlet.ShouldProcess($info.Rete, 'Salva la scansione in Censimento\scansioni')) {
        try {
            $cartella = Join-Path (Get-HDCartella -Tipo Censimento) 'scansioni'
            [void][System.IO.Directory]::CreateDirectory($cartella)
            $nomeRete = ($info.Rete -replace '[^0-9A-Za-z]+', '-').Trim('-')
            if ($nomeRete.Length -gt 60) { $nomeRete = $nomeRete.Substring(0, 60).Trim('-') }
            $copia = Join-Path $cartella ('Scansione_{0}_{1}.csv' -f $nomeRete, (Get-Date).ToString('yyyyMMdd-HHmm', $script:HDInvariante))
            $righe = @($dispositivi | Select-Object IP, Hostname, Nome, MAC, Produttore, Gruppo, Fonte | ConvertTo-Csv -NoTypeInformation -Delimiter ';')
            Write-HDFileTesto -Percorso $copia -Testo (($righe -join "`r`n") + "`r`n")
        }
        catch { Write-Warning "Copia della scansione non salvata: $($_.Exception.Message)" }
    }

    $risultato = [pscustomobject]@{
        PSTypeName  = 'HelpDesk.ScansioneRete'
        Rete        = $info.Rete
        Dispositivi = $dispositivi
        Secondi     = $secondi
        FileCsv     = $copia
        Import      = $null
    }
    if ($NelCensimento -and $dispositivi.Count -gt 0) {
        $scrivi = $PSCmdlet.ShouldProcess('censimento', "Aggiungi $($dispositivi.Count) dispositivi")
        $risultato.Import = Import-HDDispositiviNelCensimento -Dispositivi $dispositivi -Formato ("scansione integrata di " + $info.Rete) -Censimento $Censimento -Cliente $Cliente -NonChiedere:$NonChiedere -Scrivi $scrivi
    }
    $risultato
}

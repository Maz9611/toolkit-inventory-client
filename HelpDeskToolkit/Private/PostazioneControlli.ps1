# Verifica postazione - VALUTAZIONE dei dati raccolti.
# Funzioni pure: ricevono i dati e restituiscono esiti, quindi si testano anche
# senza Windows. Per aggiungere un controllo: scrivi una funzione Test-HDCtrlXxx
# e registrala in $script:HDControlli.

$script:HDControlli = [ordered]@{
    SistemaOperativo = 'Test-HDCtrlSistemaOperativo'
    Uptime           = 'Test-HDCtrlUptime'
    Riavvio          = 'Test-HDCtrlRiavvio'
    Disco            = 'Test-HDCtrlDisco'
    Memoria          = 'Test-HDCtrlMemoria'
    Aggiornamenti    = 'Test-HDCtrlAggiornamenti'
    Antivirus        = 'Test-HDCtrlProtezione'
    Firewall         = 'Test-HDCtrlFirewall'
    BitLocker        = 'Test-HDCtrlBitLocker'
    Rete             = 'Test-HDCtrlRete'
    Ora              = 'Test-HDCtrlOra'
    Eventi           = 'Test-HDCtrlEventi'
    DischiFisici     = 'Test-HDCtrlDischiFisici'
    Servizi          = 'Test-HDCtrlServizi'
    Stampanti        = 'Test-HDCtrlStampanti'
    PuntoRipristino  = 'Test-HDCtrlPuntoRipristino'
}

$script:HDIpv4 = '^\d{1,3}(\.\d{1,3}){3}$'

function New-HDEsito {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param(
        [Parameter(Mandatory)][string]$Area,
        [Parameter(Mandatory)][string]$Controllo,
        [Parameter(Mandatory)][ValidateSet('OK', 'ATTENZIONE', 'CRITICO', 'INFO', 'NON VERIFICATO', 'ERRORE')][string]$Stato,
        [string]$Dettaglio = '',
        [string]$Suggerimento = ''
    )
    [pscustomobject]@{
        PSTypeName   = 'HelpDesk.Controllo'
        Area         = $Area
        Controllo    = $Controllo
        Stato        = $Stato
        Dettaglio    = $Dettaglio
        Suggerimento = $Suggerimento
    }
}

function Get-HDMotivo {
    param($Dati, [string]$Chiave)
    if ($Dati.Errori -and $Dati.Errori[$Chiave]) { return [string]$Dati.Errori[$Chiave] }
    'Dato non disponibile su questo sistema'
}

function Get-HDStatoPeggiore {
    param([string[]]$Stati)
    foreach ($s in 'CRITICO', 'ATTENZIONE', 'ERRORE', 'NON VERIFICATO', 'INFO', 'OK') {
        if ($Stati -contains $s) { return $s }
    }
    'OK'
}

function ConvertFrom-HDDataWmi {
    # "20261015083000.000000+060" (formato data WMI) -> DateTime locale
    param($Valore)
    if ($Valore -is [datetime]) { return $Valore }
    $s = [string]$Valore
    if ($s -notmatch '^(\d{14})\.\d{6}([+-])(\d{3})$') { return $null }
    $d = [datetime]::ParseExact($Matches[1], 'yyyyMMddHHmmss', $script:HDInvariante)
    $scarto = [int]$Matches[3]
    if ($Matches[2] -eq '-') { $scarto = -$scarto }
    [datetime]::SpecifyKind($d.AddMinutes(-$scarto), [System.DateTimeKind]::Utc).ToLocalTime()
}

function Test-HDCtrlSistemaOperativo {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Sistema operativo'
    $os = $Dati.OS
    if (-not $os) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'OS') }
    $caption = ([string]$os.Caption -replace '^Microsoft\s+', '').Trim()
    $versione = ''
    if ($Dati.VersioneWindows -and $Dati.VersioneWindows.DisplayVersion) { $versione = ' ' + $Dati.VersioneWindows.DisplayVersion }
    $sup = Get-HDSupportoOS -Caption $caption
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato $sup.Stato -Dettaglio ('{0}{1} (build {2})' -f $caption, $versione, $os.BuildNumber) -Suggerimento $sup.Messaggio
}

function Test-HDCtrlUptime {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Ultimo riavvio'
    if (-not $Dati.OS -or -not $Dati.OS.LastBootUpTime) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'OS')
    }
    $avvio = [datetime]$Dati.OS.LastBootUpTime
    $acceso = $Adesso - $avvio
    $det = 'Acceso da {0} (avvio: {1})' -f (Format-HDTempoTrascorso $acceso), (Format-HDData $avvio)
    if ($acceso.TotalDays -ge $Soglie.UptimeGiorni) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Riavvia il PC prima di indagare: molti problemi spariscono. Con l''avvio rapido attivo "Arresta il sistema" non azzera il contatore, serve "Riavvia".'
    }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlRiavvio {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Riavvio in sospeso'
    if ($null -eq $Dati.RiavvioInSospeso) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'RiavvioInSospeso') }
    $motivi = @($Dati.RiavvioInSospeso | Where-Object { $_ })
    if ($motivi.Count -eq 0) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio 'Nessun riavvio in sospeso' }
    $importanti = @($motivi | Where-Object { $_ -ne 'Operazioni su file in attesa' })
    if ($importanti.Count -eq 0) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'INFO' -Dettaglio 'Operazioni su file in attesa (spesso lasciate da installazioni di programmi)' -Suggerimento 'Un riavvio le completa.'
    }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio ('Richiesto da: ' + ($motivi -join ', ')) -Suggerimento 'Riavvia prima di installare altro: aggiornamenti o driver lasciati a metà causano errori strani.'
}

function Test-HDCtrlDisco {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Spazio su disco'
    if ($null -eq $Dati.Dischi) { return New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Dischi') }
    $volumi = @($Dati.Dischi | Where-Object { $_ -and [double]$_.Size -gt 0 } | ForEach-Object { ConvertTo-HDVolume -Disco $_ -Soglie $Soglie })
    if ($volumi.Count -eq 0) { return New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio 'Nessun disco fisso trovato' }
    $det = ($volumi | ForEach-Object { '{0} {1} GB liberi su {2} GB ({3}%)' -f $_.Unita, (Format-HDNumero $_.LiberoGB), (Format-HDNumero $_.TotaleGB 0), (Format-HDNumero $_.LiberoPct 0) }) -join ' | '
    $stato = Get-HDStatoPeggiore @($volumi | ForEach-Object { $_.Stato })
    $sugg = ''
    if ($stato -ne 'OK') { $sugg = 'Poco spazio libero: usa la stima di pulizia (menu 2 o Get-StimaPulizia) per capire cosa liberare. Windows Update ha bisogno di almeno 10-20 GB.' }
    New-HDEsito -Area 'Dischi' -Controllo $c -Stato $stato -Dettaglio $det -Suggerimento $sugg
}

function Test-HDCtrlMemoria {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Memoria RAM'
    $os = $Dati.OS
    if (-not $os -or -not $os.TotalVisibleMemorySize) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'OS') }
    $totale = [double]$os.TotalVisibleMemorySize
    $totaleGB = $totale / 1MB   # il valore è in KB
    $uso = [math]::Round(($totale - [double]$os.FreePhysicalMemory) / $totale * 100)
    $det = 'Totale {0} GB, in uso {1}%' -f (Format-HDNumero $totaleGB), $uso
    $processi = @($Dati.Processi | Where-Object { $_ })
    if ($processi.Count -gt 0) {
        $det += ' | Più pesanti: ' + (($processi | Select-Object -First 5 | ForEach-Object { '{0} {1}' -f $_.Nome, (Format-HDByte ([double]$_.MemoriaMB * 1MB)) }) -join ', ')
    }
    if ($totaleGB -lt 3.5) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Meno di 4 GB di RAM: Windows 10/11 e il browser faticano. Valuta un''espansione.'
    }
    if ($uso -ge $Soglie.RamUsoPct) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Memoria quasi esaurita: chiudi i programmi più pesanti (vedi elenco) o valuta più RAM.'
    }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlAggiornamenti {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Aggiornamenti Windows'
    if ($null -eq $Dati.Qfe) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Qfe') }
    $ultimo = $null
    $kb = ''
    foreach ($q in @($Dati.Qfe | Where-Object { $_ })) {
        $d = ConvertFrom-HDDataQfe -Valore $q.InstalledOn -Adesso $Adesso
        if ($d -and (-not $ultimo -or $d -gt $ultimo)) { $ultimo = $d; $kb = $q.HotFixID }
    }
    if (-not $ultimo) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio 'Nessuna data di installazione leggibile' }
    $giorni = [int][math]::Floor(($Adesso - $ultimo).TotalDays)
    $det = 'Ultimo aggiornamento installato: {0} del {1} ({2})' -f $kb, (Format-HDData $ultimo 'dd/MM/yyyy'), (Format-HDGiorniFa $giorni)
    if ($giorni -ge $Soglie.AggiornamentiGiorniCritico) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'CRITICO' -Dettaglio $det -Suggerimento 'Da mesi non si installano aggiornamenti: controlla Windows Update (servizio, spazio su disco, errori in Impostazioni).'
    }
    if ($giorni -ge $Soglie.AggiornamentiGiorniAttenzione) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Aggiornamenti in ritardo: verifica Impostazioni > Windows Update.'
    }
    New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlProtezione {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Antivirus'
    $prodotti = $Dati.Antivirus
    $def = $Dati.Defender
    if ($null -eq $prodotti -and $null -eq $def) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Antivirus') }
    $attivi = @()
    $descr = @()
    foreach ($p in @($prodotti | Where-Object { $_ })) {
        $st = ConvertFrom-HDStatoProtezione -ProductState $p.productState
        $testo = 'non attivo'
        if ($st.Attivo) {
            $testo = 'attivo'
            if (-not $st.Aggiornato) { $testo = 'attivo, firme NON aggiornate' }
            $attivi += [pscustomobject]@{ Nome = $p.displayName; Aggiornato = $st.Aggiornato }
        }
        $descr += '{0} ({1})' -f $p.displayName, $testo
    }
    $defenderAttivo = [bool]($def -and $def.RealTimeProtectionEnabled)
    if ($def) {
        $rt = 'disattivata'
        if ($def.RealTimeProtectionEnabled) { $rt = 'attiva' }
        $descr += 'Defender: protezione in tempo reale {0}, firme di {1} giorni' -f $rt, $def.AntivirusSignatureAge
    }
    $det = $descr -join ' | '
    if ($attivi.Count -eq 0 -and -not $defenderAttivo) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'CRITICO' -Dettaglio ("Nessun antivirus attivo. $det").Trim() -Suggerimento 'Riattiva Microsoft Defender o verifica licenza e stato dell''antivirus installato prima di proseguire.'
    }
    if (@($attivi | Where-Object { -not $_.Aggiornato }).Count -gt 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Firme antivirus non aggiornate: avvia l''aggiornamento e controlla che il PC raggiunga Internet.'
    }
    if ($defenderAttivo -and [int]$def.AntivirusSignatureAge -gt $Soglie.FirmeAntivirusGiorni) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Firme di Defender vecchie: aggiornale da Sicurezza di Windows (o Update-MpSignature).'
    }
    New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlFirewall {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Firewall'
    $profili = @($Dati.Firewall | Where-Object { $_ })
    if ($null -eq $Dati.Firewall -or $profili.Count -eq 0) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Firewall') }
    $nomi = @{ Domain = 'Dominio'; Private = 'Privato'; Public = 'Pubblico' }
    $voci = foreach ($p in $profili) {
        $n = $nomi[[string]$p.Name]
        if (-not $n) { $n = [string]$p.Name }
        [pscustomobject]@{ Nome = $n; Attivo = ([string]$p.Enabled -eq 'True') }
    }
    $det = ($voci | ForEach-Object { if ($_.Attivo) { "$($_.Nome): attivo" } else { "$($_.Nome): DISATTIVATO" } }) -join ', '
    $spenti = @($voci | Where-Object { -not $_.Attivo } | ForEach-Object { $_.Nome })
    if ($spenti.Count -gt 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento ('Firewall disattivato sul profilo {0}. Con un firewall di terze parti (es. suite antivirus) può essere normale, altrimenti riattivalo.' -f ($spenti -join ', '))
    }
    New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlBitLocker {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'BitLocker'
    $volumi = @($Dati.BitLocker | Where-Object { $_ })
    if ($null -eq $Dati.BitLocker -or $volumi.Count -eq 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'BitLocker') -Suggerimento 'Avvia il toolkit come amministratore per leggere lo stato di BitLocker.'
    }
    $det = ($volumi | ForEach-Object { '{0} {1}' -f $_.MountPoint, ([string]$_.Protezione).ToLower() }) -join ', '
    if ($volumi[0].Fonte) { $det += " ($($volumi[0].Fonte))" }
    $chiave = 'Prima di aggiornare il BIOS, cambiare hardware o azzerare il TPM procurati la chiave di ripristino (account Microsoft, Active Directory/Entra ID o copia stampata) oppure sospendi la protezione (Suspend-BitLocker -RebootCount 1).'
    if (@($volumi | Where-Object { $_.Protezione -match '^Sospesa' }).Count -gt 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento ('Protezione sospesa: se non è voluto riprendila (Resume-BitLocker). ' + $chiave)
    }
    if (@($volumi | Where-Object { $_.Protezione -match '^(Attiva|Crittografia)' }).Count -gt 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'INFO' -Dettaglio $det -Suggerimento $chiave
    }
    New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'INFO' -Dettaglio $det
}

function Test-HDCtrlRete {
    param($Dati, $Soglie, [datetime]$Adesso)
    $a = 'Rete'
    if ($null -eq $Dati.Rete) { return New-HDEsito -Area $a -Controllo 'Scheda di rete' -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Rete') }
    $schede = @($Dati.Rete | Where-Object { $_ })
    $conIp = @($schede | Where-Object { @(@($_.IPAddress) -match $script:HDIpv4).Count -gt 0 })
    if ($conIp.Count -eq 0) {
        New-HDEsito -Area $a -Controllo 'Indirizzo IP' -Stato 'CRITICO' -Dettaglio 'Nessuna scheda di rete con indirizzo IPv4' -Suggerimento 'Controlla cavo o Wi-Fi e che la scheda di rete sia abilitata.'
        return
    }
    $ip = @($conIp | ForEach-Object { @($_.IPAddress) -match $script:HDIpv4 })
    $buoni = @($ip | Where-Object { $_ -notlike '169.254.*' })
    if ($buoni.Count -eq 0) {
        New-HDEsito -Area $a -Controllo 'Indirizzo IP' -Stato 'CRITICO' -Dettaglio ('Solo indirizzi automatici: {0}' -f ($ip -join ', ')) -Suggerimento 'Indirizzo 169.254.x.x = il DHCP non risponde: controlla cavo, Wi-Fi, switch o router.'
    }
    else {
        $nomiSchede = ($conIp | ForEach-Object { $_.Description }) -join ', '
        New-HDEsito -Area $a -Controllo 'Indirizzo IP' -Stato 'OK' -Dettaglio ('{0} ({1})' -f ($buoni -join ', '), $nomiSchede)
    }

    $gw = @($conIp | ForEach-Object { @($_.DefaultIPGateway) -match $script:HDIpv4 } | Select-Object -Unique)
    if ($gw.Count -eq 0) {
        New-HDEsito -Area $a -Controllo 'Gateway' -Stato 'ATTENZIONE' -Dettaglio 'Nessun gateway predefinito' -Suggerimento 'Senza gateway non si esce dalla rete locale: controlla la configurazione IP/DHCP.'
    }
    else {
        $ping = $Dati.PingGateway
        if (-not $ping) {
            New-HDEsito -Area $a -Controllo 'Gateway' -Stato 'INFO' -Dettaglio ('{0} (ping non eseguito)' -f ($gw -join ', '))
        }
        else {
            $muti = @($gw | Where-Object { $ping.ContainsKey($_) -and -not $ping[$_] })
            if ($muti.Count -gt 0) {
                New-HDEsito -Area $a -Controllo 'Gateway' -Stato 'ATTENZIONE' -Dettaglio ('{0} non risponde al ping' -f ($muti -join ', ')) -Suggerimento 'Alcuni router non rispondono al ping: se Internet funziona è normale.'
            }
            else {
                New-HDEsito -Area $a -Controllo 'Gateway' -Stato 'OK' -Dettaglio ('{0} raggiungibile' -f ($gw -join ', '))
            }
        }
    }

    $serverDns = @($conIp | ForEach-Object { @($_.DNSServerSearchOrder) } | Where-Object { $_ } | Select-Object -Unique) -join ', '
    if ($Dati.SenzaInternet) {
        New-HDEsito -Area $a -Controllo 'DNS' -Stato 'INFO' -Dettaglio "Prova saltata (modalità senza Internet). Server DNS: $serverDns"
        New-HDEsito -Area $a -Controllo 'Internet' -Stato 'INFO' -Dettaglio 'Prova saltata (modalità senza Internet)'
        return
    }
    if ($Dati.Dns -and $Dati.Dns.Ok) {
        New-HDEsito -Area $a -Controllo 'DNS' -Stato 'OK' -Dettaglio "Risolve i nomi (server: $serverDns)"
    }
    else {
        New-HDEsito -Area $a -Controllo 'DNS' -Stato 'CRITICO' -Dettaglio ("Non risolve i nomi. Server: $serverDns. " + (Get-HDMotivo $Dati 'Dns')) -Suggerimento 'Controlla i server DNS impostati (router o server di dominio) e prova: ipconfig /flushdns.'
    }
    if ($Dati.Internet -and $Dati.Internet.Ok) {
        New-HDEsito -Area $a -Controllo 'Internet' -Stato 'OK' -Dettaglio 'Connessione a Internet funzionante'
    }
    elseif ($Dati.Internet) {
        New-HDEsito -Area $a -Controllo 'Internet' -Stato 'ATTENZIONE' -Dettaglio 'Risposta inattesa dal sito di prova Microsoft' -Suggerimento 'Probabile proxy, filtro web o portale di accesso (hotspot) che intercetta le pagine.'
    }
    else {
        New-HDEsito -Area $a -Controllo 'Internet' -Stato 'CRITICO' -Dettaglio ('Internet non raggiungibile: ' + (Get-HDMotivo $Dati 'Internet')) -Suggerimento 'Verifica router/firewall/proxy. Se la rete locale funziona il problema è a monte.'
    }
}

function Test-HDCtrlOra {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Data e ora'
    $fuso = " | Fuso orario: $($Dati.FusoOrario)"
    if ($Dati.SenzaInternet) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'INFO' -Dettaglio ("Confronto con Internet saltato$fuso") }
    $i = $Dati.Internet
    if (-not $i -or -not $i.DataServerUtc) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio ("Impossibile confrontare l'ora con un server Internet$fuso")
    }
    $scarto = ([datetime]$i.OraLocaleUtc - [datetime]$i.DataServerUtc).TotalMinutes
    $assoluto = [math]::Abs($scarto)
    if ($assoluto -lt 1) { $det = 'Orologio allineato (scarto {0} secondi)' -f [math]::Round($assoluto * 60) }
    else {
        $verso = 'indietro'
        if ($scarto -gt 0) { $verso = 'avanti' }
        $det = 'Orologio {0} di {1} minuti' -f $verso, [math]::Round($assoluto)
    }
    $det += $fuso
    if ($assoluto -ge $Soglie.ScartoOrarioMinuti) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Con l''orologio sbagliato falliscono accesso al dominio, siti HTTPS, Microsoft 365 e Teams. Controlla fuso orario e sincronizzazione (Impostazioni > Data/ora > Sincronizza ora).'
    }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlEventi {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Registro eventi'
    if ($null -eq $Dati.Eventi -and $null -eq $Dati.EventiDisco) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Eventi') }
    $ev = @($Dati.Eventi | Where-Object { $_ })
    $evDisco = @($Dati.EventiDisco | Where-Object { $_ })
    $stato = 'OK'
    $problemi = @()
    $sugg = @()

    $disco = @($evDisco | Where-Object { ($_.ProviderName -eq 'disk' -and [int]$_.Id -in 7, 11, 51, 52, 153) -or ($_.ProviderName -match 'Ntfs' -and [int]$_.Id -eq 55) })
    if ($disco.Count -gt 0) {
        $stato = 'CRITICO'
        $ultimo = ($disco | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated
        $problemi += '{0} in 30 giorni (ultimo: {1})' -f (Format-HDQuantita $disco.Count 'errore del disco' 'errori del disco'), (Format-HDData $ultimo)
        $sugg += 'Possibile disco in avaria: FAI SUBITO UN BACKUP dei dati, prima di qualsiasi altra operazione.'
    }
    $bsod = @($ev | Where-Object { [int]$_.Id -eq 1001 -and $_.ProviderName -match 'WER-SystemErrorReporting|BugCheck' })
    if ($bsod.Count -gt 0) {
        $stato = Get-HDStatoPeggiore @($stato, 'ATTENZIONE')
        $problemi += Format-HDQuantita $bsod.Count 'schermata blu (BSOD)' 'schermate blu (BSOD)'
        $sugg += 'Schermate blu: controlla i dump in C:\Windows\Minidump, driver aggiornati di recente e RAM.'
    }
    $spegnimenti = @($ev | Where-Object { [int]$_.Id -eq 41 -and $_.ProviderName -match 'Kernel-Power' })
    if ($spegnimenti.Count -gt 0) {
        $stato = Get-HDStatoPeggiore @($stato, 'ATTENZIONE')
        $problemi += Format-HDQuantita $spegnimenti.Count 'spegnimento imprevisto' 'spegnimenti imprevisti'
        $sugg += 'Spegnimenti imprevisti: alimentazione, surriscaldamento o blocchi. Chiedi all''utente se il PC si spegne da solo.'
    }
    $whea = @($ev | Where-Object { $_.ProviderName -match 'WHEA' })
    if ($whea.Count -gt 0) {
        $stato = Get-HDStatoPeggiore @($stato, 'ATTENZIONE')
        $problemi += Format-HDQuantita $whea.Count 'errore hardware (WHEA)' 'errori hardware (WHEA)'
        $sugg += 'Errori hardware segnalati: verifica temperature, RAM e firmware/BIOS.'
    }
    if ($ev.Count -gt 0) {
        $frequenti = $ev | Group-Object -Property ProviderName | Sort-Object Count -Descending | Select-Object -First 3 |
            ForEach-Object { '{0} ({1})' -f ($_.Name -replace '^Microsoft-Windows-', ''), $_.Count }
        $problemi += '{0} in 7 giorni, i più frequenti: {1}' -f (Format-HDQuantita $ev.Count 'evento di errore' 'eventi di errore'), ($frequenti -join ', ')
        if ($stato -eq 'OK') { $stato = 'INFO' }
    }
    if ($problemi.Count -eq 0) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio 'Nessun errore negli ultimi 7 giorni' }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato $stato -Dettaglio ($problemi -join ' | ') -Suggerimento ($sugg -join ' ')
}

function Test-HDCtrlDischiFisici {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Salute dei dischi'
    $dischi = @($Dati.DischiFisici | Where-Object { $_ })
    if ($null -eq $Dati.DischiFisici -or $dischi.Count -eq 0) { return New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'DischiFisici') }
    $mappa = @{ '0' = 'Integro'; 'Healthy' = 'Integro'; '1' = 'Avviso'; 'Warning' = 'Avviso'; '2' = 'Guasto'; 'Unhealthy' = 'Guasto'; '5' = 'Sconosciuto'; 'Unknown' = 'Sconosciuto' }
    $voci = foreach ($d in $dischi) {
        $s = $mappa[[string]$d.HealthStatus]
        if (-not $s) { $s = [string]$d.HealthStatus }
        [pscustomobject]@{ Nome = ([string]$d.FriendlyName).Trim(); Salute = $s }
    }
    $det = ($voci | ForEach-Object { '{0}: {1}' -f $_.Nome, $_.Salute }) -join ' | '
    if (@($voci | Where-Object { $_.Salute -eq 'Guasto' }).Count -gt 0) {
        return New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'CRITICO' -Dettaglio $det -Suggerimento 'Disco guasto secondo Windows: backup immediato e sostituzione.'
    }
    if (@($voci | Where-Object { $_.Salute -eq 'Avviso' }).Count -gt 0) {
        return New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio $det -Suggerimento 'Il disco segnala anomalie (SMART): fai un backup e pianifica la sostituzione.'
    }
    New-HDEsito -Area 'Dischi' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Test-HDCtrlServizi {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Servizi essenziali'
    if ($null -eq $Dati.Servizi) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Servizi') }
    $fermi = @()
    $note = @()
    foreach ($s in @($Dati.Servizi | Where-Object { $_ })) {
        $nome = $script:HDServiziChiave[[string]$s.Name]
        if (-not $nome) { $nome = $s.DisplayName }
        if ($s.StartMode -eq 'Auto' -and $s.State -ne 'Running') { $fermi += '{0} ({1})' -f $nome, $s.Name }
        if ($s.Name -eq 'Spooler' -and $s.StartMode -eq 'Disabled') { $note += 'Spooler di stampa disabilitato: da questo PC non si può stampare' }
    }
    if ($fermi.Count -gt 0) {
        return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio ('Servizi automatici fermi: ' + ($fermi -join ', ')) -Suggerimento 'Prova ad avviarli (services.msc) e cerca nel Registro eventi il motivo dell''arresto.'
    }
    if ($note.Count -gt 0) { return New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'INFO' -Dettaglio ($note -join ' | ') }
    New-HDEsito -Area 'Sistema' -Controllo $c -Stato 'OK' -Dettaglio 'I servizi essenziali controllati sono attivi'
}

function Test-HDCtrlStampanti {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Stampanti'
    if ($null -eq $Dati.Stampanti) { return New-HDEsito -Area 'Periferiche' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'Stampanti') }
    $tutte = @($Dati.Stampanti | Where-Object { $_ })
    if ($tutte.Count -eq 0) { return New-HDEsito -Area 'Periferiche' -Controllo $c -Stato 'INFO' -Dettaglio 'Nessuna stampante installata' }
    $virtuali = '(?i)PDF|XPS|OneNote|Fax|Send To|Inoltra a'
    $reali = @($tutte | Where-Object { $_.Name -notmatch $virtuali })
    $predefinita = @($tutte | Where-Object { $_.Default }) | Select-Object -First 1
    $det = '{0} ({1})' -f (Format-HDQuantita $tutte.Count 'stampante' 'stampanti'), (Format-HDQuantita $reali.Count 'fisica' 'fisiche')
    if ($predefinita) { $det += ', predefinita: ' + $predefinita.Name }
    $stato = 'OK'
    $sugg = @()
    $coda = @($Dati.CodaStampa | Where-Object { $_ })
    if ($coda.Count -gt 0) {
        $stato = 'ATTENZIONE'
        $inErrore = @($coda | Where-Object { [string]$_.JobStatus -match '(?i)error|errore' -or [string]$_.Status -match '(?i)error|degraded' })
        $det += ' | {0} ({1} in errore)' -f (Format-HDQuantita $coda.Count 'lavoro in coda' 'lavori in coda'), $inErrore.Count
        $sugg += 'Se la coda è bloccata: annulla i lavori e riavvia lo Spooler (da amministratore: Restart-Service Spooler).'
    }
    $offline = @($reali | Where-Object { $_.WorkOffline -eq $true })
    if ($offline.Count -gt 0) {
        $stato = 'ATTENZIONE'
        $det += ' | Offline: ' + (($offline | ForEach-Object { $_.Name }) -join ', ')
        $sugg += 'Stampanti offline: controlla accensione, cavo/rete e la porta configurata.'
    }
    New-HDEsito -Area 'Periferiche' -Controllo $c -Stato $stato -Dettaglio $det -Suggerimento ($sugg -join ' ')
}

function Test-HDCtrlPuntoRipristino {
    param($Dati, $Soglie, [datetime]$Adesso)
    $c = 'Punto di ripristino'
    $crea = 'Prima di installare driver o programmi crea un punto di ripristino (da amministratore: Checkpoint-Computer -Description "Prima dell''intervento").'
    if (-not $Dati.Admin) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio 'Serve avviare il toolkit come amministratore' -Suggerimento $crea }
    if ($null -eq $Dati.PuntiRipristino) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'NON VERIFICATO' -Dettaglio (Get-HDMotivo $Dati 'PuntiRipristino') -Suggerimento $crea }
    $date = @($Dati.PuntiRipristino | Where-Object { $_ } | ForEach-Object { ConvertFrom-HDDataWmi $_.CreationTime } | Where-Object { $_ } | Sort-Object -Descending)
    if ($date.Count -eq 0) {
        return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'ATTENZIONE' -Dettaglio 'Nessun punto di ripristino disponibile' -Suggerimento ('La Protezione sistema potrebbe essere disattivata. ' + $crea)
    }
    $giorni = [int][math]::Floor(($Adesso - $date[0]).TotalDays)
    $det = 'Ultimo punto: {0} ({1}), {2} in totale' -f (Format-HDData $date[0]), (Format-HDGiorniFa $giorni), $date.Count
    if ($giorni -gt 30) { return New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'INFO' -Dettaglio $det -Suggerimento $crea }
    New-HDEsito -Area 'Sicurezza' -Controllo $c -Stato 'OK' -Dettaglio $det
}

function Invoke-HDControlli {
    # Esegue i controlli registrati; un controllo che va in errore non blocca gli altri.
    param([Parameter(Mandatory)]$Dati, [hashtable]$Soglie, [datetime]$Adesso = (Get-Date), [string[]]$Escludi)
    if (-not $Soglie) { $Soglie = (Get-HDConfig).Soglie }
    foreach ($nome in @($script:HDControlli.Keys)) {
        if ($Escludi -contains $nome) { continue }
        $funzione = $script:HDControlli[$nome]
        try {
            & $funzione -Dati $Dati -Soglie $Soglie -Adesso $Adesso
        }
        catch {
            New-HDEsito -Area 'Toolkit' -Controllo $nome -Stato 'ERRORE' -Dettaglio "Il controllo si è interrotto: $($_.Exception.Message)" -Suggerimento 'È un problema del toolkit, non del PC: annota il messaggio e correggi lo script.'
        }
    }
}

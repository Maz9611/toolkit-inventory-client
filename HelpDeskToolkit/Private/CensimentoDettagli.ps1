# Censimento: dati in più nei fogli "Dettagli" (una riga per PC) e "Software" (una riga per
# programma installato). Le 20 colonne del foglio principale non cambiano.
# Funzioni pure: conversione dei dati grezzi raccolti sul PC e modifiche da fare nei due fogli.

function New-HDColonnaFoglio {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param([string]$Chiave, [string]$Titolo, [double]$Larghezza)
    [pscustomobject]@{ Chiave = $Chiave; Titolo = $Titolo; Larghezza = $Larghezza; Sinonimi = @(); Manuale = $false; Numerica = $false; Testo = $true }
}

$script:HDFoglioDettagli = 'Dettagli'
$script:HDFoglioSoftware = 'Software'

$script:HDColonneDettagli = @(
    New-HDColonnaFoglio 'Hostname' 'Hostname' 18
    New-HDColonnaFoglio 'SerialNumber' 'Serial Number' 16
    New-HDColonnaFoglio 'Windows11' 'Pronto per Windows 11' 34
    New-HDColonnaFoglio 'TPM' 'TPM' 16
    New-HDColonnaFoglio 'SecureBoot' 'Secure Boot' 14
    New-HDColonnaFoglio 'Avvio' 'Avvio' 9
    New-HDColonnaFoglio 'DischiFisici' 'Dischi fisici' 48
    New-HDColonnaFoglio 'BitLocker' 'BitLocker' 24
    New-HDColonnaFoglio 'Bios' 'BIOS' 24
    New-HDColonnaFoglio 'Monitor' 'Monitor' 44
    New-HDColonnaFoglio 'Stampanti' 'Stampanti' 44
    New-HDColonnaFoglio 'Office' 'Office' 34
    New-HDColonnaFoglio 'AttivazioneWindows' 'Attivazione Windows' 26
    New-HDColonnaFoglio 'Antivirus' 'Antivirus' 34
    New-HDColonnaFoglio 'AccessoRemoto' 'Accesso remoto' 30
    New-HDColonnaFoglio 'Programmi' 'Programmi installati' 12
    New-HDColonnaFoglio 'Rilevato' 'Rilevato il' 17
)

$script:HDColonneSoftware = @(
    New-HDColonnaFoglio 'Hostname' 'Hostname' 18
    New-HDColonnaFoglio 'Programma' 'Programma' 46
    New-HDColonnaFoglio 'Versione' 'Versione' 18
    New-HDColonnaFoglio 'Editore' 'Editore' 30
    New-HDColonnaFoglio 'Installato' 'Installato il' 13
    New-HDColonnaFoglio 'Presente' 'Presente' 10
    New-HDColonnaFoglio 'UltimaRilevazione' 'Ultima rilevazione' 17
)

# Voci dell'elenco programmi che non sono programmi per chi fa il censimento
$script:HDSoftwareRumore = '(?i)^(KB\d+|Security Update|Update for|Hotfix|Aggiornamento (della sicurezza )?per)|Microsoft Visual C\+\+ .*(Redistributable|Runtime)|Microsoft \.NET (Runtime|Host|AppHost|Toolset|Targeting|SDK)|Microsoft \.NET Framework .*(Targeting|SDK|Multi-Targeting)|Microsoft Windows Desktop Runtime|Microsoft ASP\.NET Core|Windows Software Development Kit|Windows SDK|^vs_|Microsoft Update Health Tools|Microsoft Edge WebView2|Office 16 Click-to-Run|Microsoft Office .*(MUI|Proofing|Shared|Language Pack|Click-to-Run)|Strumenti di correzione'

# Programmi di accesso remoto da segnalare nel foglio Dettagli
$script:HDAccessoRemoto = [ordered]@{
    'TeamViewer'            = 'TeamViewer'
    'AnyDesk'               = 'AnyDesk'
    'Supremo'               = 'Supremo'
    'Splashtop'             = 'Splashtop'
    'RustDesk'              = 'RustDesk'
    'ScreenConnect'         = 'ScreenConnect|ConnectWise Control'
    'LogMeIn / GoTo'        = 'LogMeIn|GoTo Resolve|GoToAssist'
    'VNC'                   = 'UltraVNC|TightVNC|RealVNC|VNC Server|TigerVNC'
    'Chrome Remote Desktop' = 'Chrome Remote Desktop'
    'Zoho Assist'           = 'Zoho Assist'
    'Remote Utilities'      = 'Remote Utilities'
    'Radmin'                = 'Radmin'
    'DameWare'              = 'DameWare'
}

# Codici dei produttori di monitor (EDID) più comuni
$script:HDProduttoriMonitor = @{
    DEL = 'Dell'; HWP = 'HP'; HPN = 'HP'; SAM = 'Samsung'; SEC = 'Samsung'; GSM = 'LG'; LEN = 'Lenovo'; PHL = 'Philips'; AOC = 'AOC'
    ACR = 'Acer'; ASU = 'Asus'; AUS = 'Asus'; BNQ = 'BenQ'; ENC = 'Eizo'; IVM = 'Iiyama'; VSC = 'ViewSonic'; NEC = 'NEC'; FUS = 'Fujitsu'
    MSI = 'MSI'; HKC = 'HKC'; SNY = 'Sony'; MEI = 'Panasonic'; APP = 'Apple'; HSD = 'Hannspree'
}
# Pannelli dei portatili (senza nome commerciale)
$script:HDPannelliIntegrati = '^(LGD|AUO|BOE|SDC|CMN|SHP|IVO|NCP|CSO|TMA|KDB|LEN|INX|BOE|MS_)$'

#region Pronto per Windows 11

function Test-HDCpuWindows11 {
    # Processore nell'elenco di Windows 11? 'Sì', 'No' oppure '' (da verificare a mano)
    param([string]$Nome)
    $n = ([string]$Nome -replace '\((R|TM|C)\)', '' -replace '\s+', ' ').Trim()
    if (-not $n) { return '' }
    if ($n -match '(?i)\bcore ultra\b|\bcore [3579] \d{3}[a-z]*\b|\bryzen ai\b|\bsnapdragon x\b') { return 'Sì' }
    if ($n -match '(?i)\b(i3-)?N(50|95|97|100|150|200|250|300|305|355)\b') { return 'Sì' }
    if ($n -match '(?i)\bi[3579]-(\d{3,5})') {
        $cifre = $Matches[1]
        if ($cifre.Length -eq 3) { return 'No' }
        $generazione = [int]$cifre.Substring(0, 1)
        if ($cifre.Length -eq 5 -or ($cifre.Length -eq 4 -and $cifre.StartsWith('1'))) { $generazione = [int]$cifre.Substring(0, 2) }
        if ($generazione -ge 8) { return 'Sì' }
        return 'No'
    }
    if ($n -match '(?i)\bryzen (threadripper )?(pro )?[3579] (pro )?(\d)(\d{3})([a-z]*)') {
        $serie = [int]$Matches[4]
        if ($serie -ge 3) { return 'Sì' }
        if ($serie -eq 1) { return 'No' }
        if ($Matches[6] -match '(?i)^x?$') { return 'Sì' }
        return ''
    }
    if ($n -match '(?i)\bathlon (gold |silver )?3\d{3}') { return 'Sì' }
    if ($n -match '(?i)core ?2|core duo|pentium (dual|d\b|4|m\b)|celeron (d|m)\b|\batom\b|\bfx-\d|\bamd [ae]\d?-\d|phenom|turion|sempron') { return 'No' }
    ''
}

function ConvertFrom-HDTpm {
    # Versione del TPM: '2.0', '1.2', 'Assente' oppure '' se non si è potuto leggere
    param($Wmi, $Dispositivi)
    if ($Wmi -and $Wmi.SpecVersion) {
        $versione = (([string]$Wmi.SpecVersion) -split ',')[0].Trim()
        if ($Wmi.IsEnabled_InitialValue -eq $false) { return "$versione (disattivato)" }
        return $versione
    }
    if ($Dispositivi) {
        foreach ($nome in @($Dispositivi.Nomi)) {
            if ([string]$nome -match '(?i)trusted platform module\s*(\d\.\d)') { return $Matches[1] }
            if ([string]$nome -match '(?i)\btpm\b|trusted platform') { return 'presente' }
        }
        return 'Assente'
    }
    ''
}

function ConvertFrom-HDAvvio {
    param($Valore)
    switch ([string]$Valore) {
        '1' { 'BIOS' }
        '2' { 'UEFI' }
        default { '' }
    }
}

function ConvertFrom-HDSecureBoot {
    param($Stato, [string]$Avvio)
    if ($Avvio -eq 'BIOS') { return 'Non supportato' }
    if ($null -eq $Stato) { return '' }
    if ([string]$Stato.Attivo -eq '1') { return 'Attivo' }
    'Disattivato'
}

function Get-HDProntoWindows11 {
    # "Già Windows 11", "Sì", "No: <motivi>" oppure "Da verificare: <cosa>"
    param([string]$SistemaOperativo, [bool]$Server, [string]$Tpm, [string]$SecureBoot, [string]$Avvio, [string]$Cpu, [double]$RamGB, [double]$DiscoGB)
    if ($SistemaOperativo -match 'Windows 11') { return 'Già Windows 11' }
    if ($Server -or $SistemaOperativo -match '(?i)server') { return '' }
    $problemi = @()
    $dubbi = @()
    if ($Tpm -match '^2\.0$') { }
    elseif ($Tpm -match '^2\.0 \(disattivato\)') { $problemi += 'TPM 2.0 disattivato (va attivato nel BIOS)' }
    elseif ($Tpm -match '^1\.') { $problemi += "TPM $Tpm (serve 2.0)" }
    elseif ($Tpm -eq 'Assente') { $problemi += 'TPM assente' }
    else { $dubbi += 'versione del TPM (avvia come amministratore)' }
    if ($Avvio -eq 'BIOS') { $problemi += 'avvio BIOS legacy (serve UEFI)' }
    elseif (-not $Avvio) { $dubbi += 'avvio UEFI' }
    if ($Avvio -eq 'UEFI' -and $SecureBoot -eq 'Non supportato') { $problemi += 'Secure Boot non supportato' }
    if ($Cpu -eq 'No') { $problemi += 'processore non supportato' }
    elseif (-not $Cpu) { $dubbi += 'processore' }
    if ($RamGB -gt 0 -and $RamGB -lt 3.8) { $problemi += 'RAM sotto 4 GB' }
    if ($DiscoGB -gt 0 -and $DiscoGB -lt 60) { $problemi += 'disco sotto 64 GB' }
    if ($problemi.Count -gt 0) { return 'No: ' + ($problemi -join ', ') }
    if ($dubbi.Count -gt 0) { return 'Da verificare: ' + ($dubbi -join ', ') }
    if ($SecureBoot -eq 'Disattivato') { return 'Sì (prima attiva Secure Boot nel BIOS)' }
    'Sì'
}

#endregion

#region Hardware

function Format-HDDischiFisici {
    # "Samsung SSD 980 500GB (SSD NVMe, 466 GB, OK); ..." senza i dischi USB
    param($Dischi, $DischiWin32)
    $tipi = @{ '3' = 'HDD'; '4' = 'SSD'; '5' = 'SCM' }
    $bus = @{ '1' = 'SCSI'; '3' = 'ATA'; '8' = 'RAID'; '10' = 'SAS'; '11' = 'SATA'; '12' = 'SD'; '17' = 'NVMe' }
    $salute = @{ '0' = 'OK'; '1' = 'ATTENZIONE'; '2' = 'GUASTO' }
    $voci = foreach ($d in @($Dischi | Where-Object { $_ })) {
        if ([string]$d.BusType -eq '7') { continue }
        $tipo = @(@($tipi[[string]$d.MediaType], $bus[[string]$d.BusType]) | Where-Object { $_ }) -join ' '
        $parti = @()
        if ($tipo) { $parti += $tipo }
        if ([double]$d.Size -gt 0) { $parti += ('{0} GB' -f [math]::Round([double]$d.Size / 1GB)) }
        $stato = $salute[[string]$d.HealthStatus]
        if ($stato) { $parti += $stato }
        $nome = ([string]$d.FriendlyName).Trim()
        if ($parti.Count -gt 0) { '{0} ({1})' -f $nome, ($parti -join ', ') } else { $nome }
    }
    if (@($voci).Count -eq 0) {
        $voci = foreach ($d in @($DischiWin32 | Where-Object { $_ })) {
            if ([string]$d.InterfaceType -eq 'USB' -or [string]$d.MediaType -match '(?i)removable|esterno|external') { continue }
            $parti = @()
            if ([string]$d.InterfaceType) { $parti += [string]$d.InterfaceType }
            if ([double]$d.Size -gt 0) { $parti += ('{0} GB' -f [math]::Round([double]$d.Size / 1GB)) }
            if ([string]$d.Status -and [string]$d.Status -ne 'OK') { $parti += ('stato ' + $d.Status) }
            '{0} ({1})' -f ([string]$d.Model).Trim(), ($parti -join ', ')
        }
    }
    @($voci) -join '; '
}

function Format-HDBitLocker {
    # "C: attiva; D: non attiva"
    param($Admin, $Shell)
    $voci = @()
    if ($Admin) {
        foreach ($v in @($Admin | Where-Object { $_.DriveLetter } | Sort-Object DriveLetter)) {
            $stato = switch ([string]$v.ProtectionStatus) { '1' { 'attiva' } '0' { 'non attiva' } default { 'sconosciuta' } }
            $voci += '{0} {1}' -f $v.DriveLetter, $stato
        }
    }
    elseif ($Shell) {
        foreach ($v in @($Shell | Where-Object { $_.Unita } | Sort-Object Unita)) {
            $voci += '{0} {1}' -f $v.Unita, (ConvertFrom-HDBitLockerShell $v.Valore).ToLowerInvariant()
        }
    }
    $voci -join '; '
}

function Format-HDFirmware {
    # "1.15.0 (12/03/2022)"
    param($Bios)
    if (-not $Bios) { return '' }
    $versione = ([string]$Bios.SMBIOSBIOSVersion).Trim()
    $data = $null
    if ($Bios.ReleaseDate -is [datetime]) { $data = $Bios.ReleaseDate }
    elseif ([string]$Bios.ReleaseDate -match '^(\d{4})(\d{2})(\d{2})') { $data = New-Object datetime ([int]$Matches[1]), ([int]$Matches[2]), ([int]$Matches[3]) }
    if ($data -and $data.Year -gt 1990) {
        if ($versione) { return '{0} ({1})' -f $versione, (Format-HDData $data 'dd/MM/yyyy') }
        return Format-HDData $data 'dd/MM/yyyy'
    }
    $versione
}

function Format-HDMonitor {
    # "Dell P2419H (S/N CFV9N93, 2020); Schermo integrato"
    param($Monitor)
    $voci = foreach ($m in @($Monitor | Where-Object { $_ })) {
        $codice = ([string]$m.Produttore).Trim().ToUpperInvariant()
        $nome = ([string]$m.Nome).Trim()
        if (-not $nome) {
            if ($codice -match $script:HDPannelliIntegrati) { 'Schermo integrato'; continue }
            $nome = 'Monitor'
        }
        $marca = $script:HDProduttoriMonitor[$codice]
        if ($marca -and $nome -notmatch [regex]::Escape($marca)) { $nome = "$marca $nome" }
        elseif (-not $marca -and $codice -and $nome -eq 'Monitor') { $nome = "Monitor $codice" }
        $info = @()
        $seriale = Get-HDNumeroSerie $m.Seriale
        if ($seriale) { $info += "S/N $seriale" }
        if ([int]$m.Anno -gt 1990) { $info += [string]$m.Anno }
        if ($info.Count -gt 0) { '{0} ({1})' -f $nome, ($info -join ', ') } else { $nome }
    }
    @($voci) -join '; '
}

function Format-HDStampanti {
    # "HP LaserJet M404 [192.168.1.50] (predefinita); ..." senza stampanti virtuali (PDF, XPS, OneNote, fax)
    param($Stampanti)
    $virtuali = '(?i)pdf|xps|onenote|fax|send to|invia a|snagit|microsoft print|nul:|portprompt'
    $voci = foreach ($s in @($Stampanti | Where-Object { $_ -and $_.Name })) {
        if ("$($s.Name) $($s.DriverName) $($s.PortName)" -match $virtuali) { continue }
        $voce = ([string]$s.Name).Trim()
        if ([string]$s.PortName -match '(\d{1,3}(\.\d{1,3}){3})') { $voce += " [$($Matches[1])]" }
        if ($s.Default) { $voce += ' (predefinita)' }
        $voce
    }
    @($voci) -join '; '
}

#endregion

#region Software e licenze

function ConvertFrom-HDDataInstallazione {
    # "20240315" -> "15/03/2024"
    param([string]$Testo)
    if ($Testo -match '^(\d{4})(\d{2})(\d{2})$') { return '{0}/{1}/{2}' -f $Matches[3], $Matches[2], $Matches[1] }
    ''
}

function Select-HDProgrammi {
    # Voci del registro "Installazione applicazioni" -> programmi veri, senza doppioni, in ordine alfabetico
    param($Programmi)
    $tenuti = @{}
    foreach ($p in @($Programmi | Where-Object { $_ -and $_.DisplayName })) {
        $nome = ([string]$p.DisplayName).Trim()
        if ([string]$p.SystemComponent -eq '1' -or $p.ParentKeyName) { continue }
        if ([string]$p.ReleaseType -match '(?i)update|hotfix|service pack') { continue }
        if ($nome -match $script:HDSoftwareRumore) { continue }
        $chiave = ConvertTo-HDChiave $nome
        $voce = [pscustomobject]@{
            Nome       = $nome
            Versione   = ([string]$p.DisplayVersion).Trim()
            Editore    = ([string]$p.Publisher).Trim()
            Installato = (ConvertFrom-HDDataInstallazione ([string]$p.InstallDate))
        }
        if (-not $tenuti.ContainsKey($chiave) -or ($voce.Versione -and -not $tenuti[$chiave].Versione)) { $tenuti[$chiave] = $voce }
    }
    @($tenuti.Values | Sort-Object Nome)
}

function Get-HDOfficeInstallato {
    # "Microsoft 365 Apps for business (16.0.17126)" oppure "Office Home and Business 2019"
    param($Programmi, $ClickToRun)
    $nomi = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($Programmi | Where-Object { $_ -and $_.DisplayName })) {
        $n = ([string]$p.DisplayName).Trim()
        if ($n -notmatch '(?i)^Microsoft (Office|365)\b') { continue }
        if ($n -match '(?i)MUI|Proofing|Shared|Language|Click-to-Run|Correzione|Viewer|Access database engine|Web Components|Live Meeting|Interop|Primary Interop|Runtime|Tools') { continue }
        $n = ($n -replace '(?i)\s+-\s+[a-z]{2}-[a-z]{2}$', '' -replace '^Microsoft Office ', 'Office ').Trim()
        if (-not $nomi.Contains($n)) { $nomi.Add($n) }
    }
    if ($nomi.Count -eq 0 -and $ClickToRun -and $ClickToRun.ProductReleaseIds) {
        foreach ($id in ([string]$ClickToRun.ProductReleaseIds -split ',')) { if ($id.Trim()) { $nomi.Add($id.Trim()) } }
    }
    $testo = @($nomi) -join '; '
    if ($testo -and $ClickToRun -and [string]$ClickToRun.VersionToReport -match '^(\d+\.\d+\.\d+)') { $testo += " ($($Matches[1]))" }
    $testo
}

function Get-HDAttivazioneWindows {
    # "Attivato (OEM, chiave nel BIOS)", "Non attivato", ...
    param($Licenze)
    $l = @($Licenze | Where-Object { $_ } | Sort-Object { [int]($_.LicenseStatus -ne 1) }) | Select-Object -First 1
    if (-not $l) { return '' }
    $stato = switch ([string]$l.LicenseStatus) {
        '1' { 'Attivato' }
        '0' { 'Non attivato' }
        '5' { 'Non attivato (notifiche)' }
        default { 'Non attivato (periodo di prova)' }
    }
    $canale = ''
    if ([string]$l.Description -match '(?i)\b(OEM_DM|OEM_COA_NSLP|OEM_COA_SLP|OEM\w*|RETAIL|VOLUME_KMSCLIENT|VOLUME_MAK|VOLUME\w*)\s+channel') {
        $canale = switch -Regex ($Matches[1].ToUpperInvariant()) {
            '^OEM_DM' { 'OEM, chiave nel BIOS'; break }
            '^OEM' { 'OEM'; break }
            '^RETAIL' { 'Retail'; break }
            'KMS' { 'Volume KMS'; break }
            'MAK' { 'Volume MAK'; break }
            default { 'Volume' }
        }
    }
    if ($canale) { return '{0} ({1})' -f $stato, $canale }
    $stato
}

function Get-HDAntivirusInstallati {
    # "Windows Defender (attivo, aggiornato); Cynet (attivo)"
    param($Antivirus)
    $visti = @{}
    $voci = foreach ($a in @($Antivirus | Where-Object { $_ -and $_.displayName })) {
        $nome = ([string]$a.displayName).Trim()
        if ($visti.ContainsKey($nome)) { continue }
        $visti[$nome] = $true
        $stato = ConvertFrom-HDStatoProtezione -ProductState ([int64]$a.productState)
        $parti = @()
        if ($stato.Attivo) { $parti += 'attivo' } else { $parti += 'non attivo' }
        if ($stato.Aggiornato) { $parti += 'aggiornato' } else { $parti += 'firme non aggiornate' }
        '{0} ({1})' -f $nome, ($parti -join ', ')
    }
    @($voci) -join '; '
}

function Get-HDAccessoRemoto {
    # Programmi di accesso remoto installati e Desktop remoto di Windows
    param($Programmi, $DesktopRemoto)
    $voci = New-Object System.Collections.Generic.List[string]
    $nomi = @(@($Programmi | Where-Object { $_ -and $_.DisplayName }) | ForEach-Object { [string]$_.DisplayName })
    foreach ($etichetta in $script:HDAccessoRemoto.Keys) {
        if (@($nomi | Where-Object { $_ -match "(?i)$($script:HDAccessoRemoto[$etichetta])" }).Count -gt 0) { $voci.Add($etichetta) }
    }
    if ([string]$DesktopRemoto -eq '0') { $voci.Add('Desktop remoto di Windows attivo') }
    @($voci) -join '; '
}

#endregion

function ConvertTo-HDDettagliPC {
    # Dati grezzi -> valori del foglio Dettagli (hashtable) ed elenco per il foglio Software
    param([Parameter(Mandatory)]$Grezzo, [string]$SistemaOperativo, [bool]$Server, [string]$Processore)
    $tpm = ConvertFrom-HDTpm -Wmi $Grezzo.Tpm -Dispositivi $Grezzo.TpmDispositivi
    $avvio = ConvertFrom-HDAvvio $Grezzo.Avvio
    $secureBoot = ConvertFrom-HDSecureBoot -Stato $Grezzo.SecureBoot -Avvio $avvio
    $ramGB = 0
    if ($Grezzo.ComputerSystem -and $Grezzo.ComputerSystem.TotalPhysicalMemory) { $ramGB = [double]$Grezzo.ComputerSystem.TotalPhysicalMemory / 1GB }
    $somma = 0
    foreach ($m in @($Grezzo.Memoria | Where-Object { $_ })) { $somma += [double]$m.Capacity }
    if ($somma -gt 0) { $ramGB = $somma / 1GB }
    $discoGB = 0
    $sistema = @($Grezzo.Volumi | Where-Object { $_ -and [string]$_.DeviceID -eq [string]$Grezzo.UnitaSistema }) | Select-Object -First 1
    if ($sistema) { $discoGB = [double]$sistema.Size / 1GB }
    $nomeCpu = $Processore
    if (-not $nomeCpu) { $nomeCpu = [string](@($Grezzo.Processori | Where-Object { $_ }) | Select-Object -First 1).Name }

    $programmi = $null
    if ($null -ne $Grezzo.Programmi) { $programmi = Select-HDProgrammi $Grezzo.Programmi }
    $valori = [ordered]@{
        Windows11          = (Get-HDProntoWindows11 -SistemaOperativo $SistemaOperativo -Server $Server -Tpm $tpm -SecureBoot $secureBoot -Avvio $avvio -Cpu (Test-HDCpuWindows11 $nomeCpu) -RamGB $ramGB -DiscoGB $discoGB)
        TPM                = $tpm
        SecureBoot         = $secureBoot
        Avvio              = $avvio
        DischiFisici       = (Format-HDDischiFisici -Dischi $Grezzo.DischiFisici -DischiWin32 $Grezzo.DischiWin32)
        BitLocker          = (Format-HDBitLocker -Admin $Grezzo.BitLockerAdmin -Shell $Grezzo.BitLockerShell)
        Bios               = (Format-HDFirmware $Grezzo.Bios)
        Monitor            = (Format-HDMonitor $Grezzo.Monitor)
        Stampanti          = (Format-HDStampanti $Grezzo.Stampanti)
        Office             = (Get-HDOfficeInstallato -Programmi $Grezzo.Programmi -ClickToRun $Grezzo.OfficeC2R)
        AttivazioneWindows = (Get-HDAttivazioneWindows $Grezzo.Licenza)
        Antivirus          = (Get-HDAntivirusInstallati $Grezzo.Antivirus)
        AccessoRemoto      = (Get-HDAccessoRemoto -Programmi $Grezzo.Programmi -DesktopRemoto $Grezzo.DesktopRemoto)
        Programmi          = ''
    }
    if ($null -ne $programmi) {
        $valori.Programmi = [string]@($programmi).Count
        if (-not $valori.Office) { $valori.Office = 'Non installato' }
        if (-not $valori.AccessoRemoto) { $valori.AccessoRemoto = 'Nessuno' }
    }
    $dettagli = @{}
    foreach ($k in $valori.Keys) { $dettagli[$k] = [string]$valori[$k] }
    [pscustomobject]@{ Dettagli = $dettagli; Software = $programmi }
}

#region Fogli Dettagli e Software nell'Excel

function ConvertTo-HDTabellaFoglio {
    # Foglio con nome e colonne fissi (Dettagli, Software) -> righe. Esiste = $false se il foglio manca
    # (verrà creato alla scrittura); Conflitto = $true se c'è un foglio con quel nome ma con altro contenuto.
    param([object[]]$Fogli, [Parameter(Mandatory)][string]$Nome, [Parameter(Mandatory)][object[]]$Colonne)
    $mappaTitoli = @{}
    foreach ($c in $Colonne) { $mappaTitoli[(ConvertTo-HDChiave $c.Titolo)] = $c.Chiave }
    $tabella = [pscustomobject]@{
        Nome = $Nome; Definizione = $Colonne; Esiste = $false; Conflitto = $false; ScriviIntestazione = $false
        ParteFoglio = $null; IndiceFoglio = -1; RigaIntestazione = 1
        Colonne = @{}; Righe = (New-Object System.Collections.Generic.List[object]); UltimaRiga = 1
    }
    for ($i = 0; $i -lt $Colonne.Count; $i++) { $tabella.Colonne[$Colonne[$i].Chiave] = $i + 1 }
    $f = @($Fogli | Where-Object { $_ -and $_.Nome -eq $Nome }) | Select-Object -First 1
    if (-not $f) { return $tabella }
    $tabella.Esiste = $true
    $tabella.ParteFoglio = $f.Parte
    $tabella.IndiceFoglio = $f.Indice
    $intestazione = $null
    foreach ($n in @($f.Righe.Keys | Select-Object -First 15)) {
        $mappa = @{}
        foreach ($col in $f.Righe[$n].Keys) {
            $k = $mappaTitoli[(ConvertTo-HDChiave ([string]$f.Righe[$n][$col]))]
            if ($k -and -not $mappa.ContainsKey($k)) { $mappa[$k] = [int]$col }
        }
        if ($mappa.ContainsKey('Hostname') -and $mappa.Count -ge 2) { $intestazione = @{ Riga = [int]$n; Colonne = $mappa }; break }
    }
    if (-not $intestazione) {
        # foglio vuoto con quel nome: si scrive l'intestazione; con altro contenuto non si tocca
        if ($f.Righe.Count -gt 0) { $tabella.Conflitto = $true } else { $tabella.ScriviIntestazione = $true }
        return $tabella
    }
    $tabella.RigaIntestazione = $intestazione.Riga
    $tabella.Colonne = $intestazione.Colonne
    $tabella.UltimaRiga = $intestazione.Riga
    foreach ($n in $f.Righe.Keys) {
        if ($n -le $intestazione.Riga) { continue }
        if ($n -gt $tabella.UltimaRiga) { $tabella.UltimaRiga = [int]$n }
        $valori = @{}
        foreach ($k in $intestazione.Colonne.Keys) {
            $v = $f.Righe[$n][$intestazione.Colonne[$k]]
            if ($null -ne $v -and ([string]$v).Trim() -ne '') { $valori[$k] = ([string]$v).Trim() }
        }
        if ($valori.Count -gt 0) { $tabella.Righe.Add([pscustomobject]@{ Numero = [int]$n; Valori = $valori }) }
    }
    $tabella
}

function Add-HDDettagliATabella {
    # Riga del PC nel foglio Dettagli: stesso numero di serie, altrimenti stesso hostname
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][hashtable]$Valori, [Parameter(Mandatory)][hashtable]$Modifiche)
    $s = Get-HDNumeroSerie $Valori['SerialNumber']
    $h = ConvertTo-HDNomeHost $Valori['Hostname']
    $riga = $null
    if ($s) { $riga = $Tabella.Righe | Where-Object { (Get-HDNumeroSerie $_.Valori['SerialNumber']) -eq $s } | Select-Object -First 1 }
    if (-not $riga -and $h) {
        $riga = $Tabella.Righe | Where-Object {
            (ConvertTo-HDNomeHost $_.Valori['Hostname']) -eq $h -and -not ($s -and (Get-HDNumeroSerie $_.Valori['SerialNumber']))
        } | Select-Object -First 1
    }
    $azione = 'Aggiornato'
    if (-not $riga) { $riga = New-HDRigaTabella $Tabella; $azione = 'Aggiunto' }
    $nuovi = @{}
    foreach ($k in $Valori.Keys) { if ([string]$Valori[$k] -ne '') { $nuovi[$k] = [string]$Valori[$k] } }
    $diff = Get-HDDifferenze -Esistenti $riga.Valori -Nuovi $nuovi
    Add-HDRigaInTabella -Tabella $Tabella -Modifiche $Modifiche -Riga $riga -Diff $diff
    [pscustomobject]@{ Riga = $riga.Numero; Azione = $azione }
}

function Add-HDSoftwareATabella {
    # Programmi del PC nel foglio Software: aggiorna quelli già presenti, aggiunge i nuovi e segna
    # "Presente = No" quelli che non ci sono più (le righe non si cancellano mai).
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][string]$Hostname, [object[]]$Programmi = @(), [string]$Data, [Parameter(Mandatory)][hashtable]$Modifiche)
    $h = ConvertTo-HDNomeHost $Hostname
    $esistenti = @{}
    foreach ($r in $Tabella.Righe) {
        if ((ConvertTo-HDNomeHost $r.Valori['Hostname']) -ne $h) { continue }
        $k = ConvertTo-HDChiave $r.Valori['Programma']
        if ($k -and -not $esistenti.ContainsKey($k)) { $esistenti[$k] = $r }
    }
    $visti = @{}
    $aggiunti = 0
    foreach ($p in @($Programmi | Where-Object { $_ -and $_.Nome })) {
        $k = ConvertTo-HDChiave $p.Nome
        if ($visti.ContainsKey($k)) { continue }
        $visti[$k] = $true
        $riga = $esistenti[$k]
        if (-not $riga) { $riga = New-HDRigaTabella $Tabella; $aggiunti++ }
        $nuovi = @{ Hostname = $Hostname; Programma = [string]$p.Nome; Presente = 'Sì' }
        if ($Data) { $nuovi['UltimaRilevazione'] = $Data }
        foreach ($coppia in @(@('Versione', $p.Versione), @('Editore', $p.Editore), @('Installato', $p.Installato))) {
            if ([string]$coppia[1]) { $nuovi[$coppia[0]] = [string]$coppia[1] }
        }
        $diff = Get-HDDifferenze -Esistenti $riga.Valori -Nuovi $nuovi
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $Modifiche -Riga $riga -Diff $diff
    }
    $rimossi = 0
    foreach ($k in $esistenti.Keys) {
        if ($visti.ContainsKey($k) -or $esistenti[$k].Valori['Presente'] -eq 'No') { continue }
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $Modifiche -Riga $esistenti[$k] -Diff @{ Presente = 'No' }
        $rimossi++
    }
    [pscustomobject]@{ Aggiunti = $aggiunti; Rimossi = $rimossi; Totale = $visti.Count }
}

#endregion

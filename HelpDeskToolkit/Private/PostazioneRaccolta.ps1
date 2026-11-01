# Verifica postazione - RACCOLTA dei dati (solo Windows, sola lettura).
# Ogni passo è isolato: se uno fallisce (permessi, componente assente) gli altri
# proseguono e il motivo finisce in $dati.Errori.

$script:HDServiziChiave = [ordered]@{
    'Spooler'           = 'Spooler di stampa'
    'Dnscache'          = 'Client DNS'
    'Dhcp'              = 'Client DHCP'
    'LanmanWorkstation' = 'Workstation (cartelle di rete)'
    'LanmanServer'      = 'Server (condivisioni)'
    'EventLog'          = 'Registro eventi'
    'Winmgmt'           = 'WMI'
    'Schedule'          = 'Utilità di pianificazione'
    'CryptSvc'          = 'Servizi di crittografia'
    'mpssvc'            = 'Windows Defender Firewall'
    'WinDefend'         = 'Antivirus Microsoft Defender'
    'AudioSrv'          = 'Audio di Windows'
    'W32Time'           = 'Ora di Windows'
    'Netlogon'          = 'Accesso rete (dominio)'
}

function Get-HDMotiviRiavvio {
    $motivi = @()
    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $motivi += 'Componenti di Windows (CBS)' }
    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $motivi += 'Windows Update' }
    $sm = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($sm -and $sm.PendingFileRenameOperations) { $motivi += 'Operazioni su file in attesa' }
    $attivo = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction SilentlyContinue).ComputerName
    $nuovo = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction SilentlyContinue).ComputerName
    if ($attivo -and $nuovo -and $attivo -ne $nuovo) { $motivi += "Cambio nome del PC ($attivo -> $nuovo)" }
    $motivi
}

function ConvertFrom-HDBitLockerShell {
    # Valore di System.Volume.BitLockerProtection letto senza diritti di amministratore
    param($Valore)
    switch ([string]$Valore) {
        '1' { 'Attiva' }
        '2' { 'Non attiva' }
        '3' { 'Crittografia in corso' }
        '4' { 'Decrittografia in corso' }
        '5' { 'Sospesa' }
        '6' { 'Attiva (unità bloccata)' }
        '8' { 'In attesa di attivazione' }
        '' { 'Non disponibile' }
        default { "Sconosciuta (codice $Valore)" }
    }
}

function Get-HDStatoBitLocker {
    param([bool]$Admin)
    if ($Admin -and (Get-Command -Name Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        $volumi = @(Get-BitLockerVolume -ErrorAction SilentlyContinue | Where-Object { [string]$_.MountPoint -match '^[A-Za-z]:' })
        if ($volumi.Count -gt 0) {
            foreach ($v in $volumi) {
                $prot = 'Non attiva'
                if ([string]$v.ProtectionStatus -eq 'On') { $prot = 'Attiva' }
                elseif ([string]$v.VolumeStatus -match 'EncryptionInProgress') { $prot = 'Crittografia in corso' }
                elseif ([string]$v.VolumeStatus -match 'FullyEncrypted') { $prot = 'Sospesa o in attesa di attivazione' }
                [pscustomobject]@{ MountPoint = [string]$v.MountPoint; Protezione = $prot; Fonte = '' }
            }
            return
        }
    }
    # Senza diritti di amministratore: proprietà della shell di Windows (solo unità di sistema)
    $shell = New-Object -ComObject Shell.Application
    $valore = $shell.NameSpace($env:SystemDrive).Self.ExtendedProperty('System.Volume.BitLockerProtection')
    [pscustomobject]@{
        MountPoint = $env:SystemDrive
        Protezione = (ConvertFrom-HDBitLockerShell $valore)
        Fonte      = 'stima senza diritti di amministratore'
    }
}

function Test-HDPingGateway {
    param($Schede)
    $esiti = @{}
    $gateway = @(@($Schede) | Where-Object { $_ } | ForEach-Object { @($_.DefaultIPGateway) } |
            Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } | Select-Object -Unique)
    foreach ($g in $gateway) {
        try { $esiti[$g] = ((New-Object System.Net.NetworkInformation.Ping).Send($g, 1000).Status -eq 'Success') }
        catch { $esiti[$g] = $false }
    }
    $esiti
}

function ConvertFrom-HDDataHttp {
    # Intestazione HTTP "Date" (RFC 1123, sempre in UTC)
    param([string]$Valore)
    if (-not $Valore) { return $null }
    $d = [datetime]::MinValue
    $stile = [System.Globalization.DateTimeStyles]'AdjustToUniversal, AssumeUniversal'
    if ([datetime]::TryParseExact($Valore.Trim(), 'r', $script:HDInvariante, $stile, [ref]$d)) { return $d }
    $null
}

function Test-HDInternet {
    # Stessa pagina che usa Windows per decidere se c'è Internet. La data della
    # risposta serve anche a controllare l'orologio del PC.
    $richiesta = [System.Net.WebRequest]::Create('http://www.msftconnecttest.com/connecttest.txt')
    $richiesta.Timeout = 5000
    # proxy aziendali con autenticazione Windows
    if ($richiesta.Proxy) { $richiesta.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
    $risposta = $richiesta.GetResponse()
    try {
        $data = $risposta.Headers['Date']
        $lettore = New-Object System.IO.StreamReader($risposta.GetResponseStream())
        $corpo = $lettore.ReadToEnd()
        $lettore.Close()
    }
    finally { $risposta.Close() }
    [pscustomobject]@{
        Ok            = (([string]$corpo).Trim() -eq 'Microsoft Connect Test')
        DataServerUtc = (ConvertFrom-HDDataHttp $data)
        OraLocaleUtc  = [datetime]::UtcNow
    }
}

function Get-HDEventi {
    param([Parameter(Mandatory)][hashtable]$Filtro, [int]$Massimo = 500)
    try {
        Get-WinEvent -FilterHashtable $Filtro -MaxEvents $Massimo -ErrorAction Stop |
            Select-Object ProviderName, Id, Level, TimeCreated
    }
    catch {
        if ($_.FullyQualifiedErrorId -match 'NoMatchingEventsFound') { return }
        throw
    }
}

function Get-HDDatiPostazione {
    [CmdletBinding()]
    param([switch]$SenzaInternet)
    $dati = @{
        Errori        = @{}
        Admin         = (Test-HDAmministratore)
        SenzaInternet = [bool]$SenzaInternet
        Raccolto      = (Get-Date)
        NomePC        = (Get-HDNomePC)
        Utente        = ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
        FusoOrario    = [System.TimeZoneInfo]::Local.DisplayName
    }
    $filtroServizi = (@($script:HDServiziChiave.Keys) | ForEach-Object { "Name='$_'" }) -join ' OR '

    # NB: ", @(...)" restituisce sempre un elenco (anche vuoto): così "nessun elemento"
    # (elenco vuoto) si distingue da "dato non leggibile" ($null).
    $passi = [ordered]@{
        OS               = { Get-HDCim -Classe Win32_OperatingSystem | Select-Object -First 1 Caption, Version, BuildNumber, LastBootUpTime, FreePhysicalMemory, TotalVisibleMemorySize }
        CS               = { Get-HDCim -Classe Win32_ComputerSystem | Select-Object -First 1 Name, Domain, PartOfDomain, Manufacturer, Model, UserName }
        Bios             = { Get-HDCim -Classe Win32_BIOS | Select-Object -First 1 SerialNumber }
        VersioneWindows  = { Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop | Select-Object DisplayVersion, CurrentBuild, UBR }
        Dischi           = { , @(Get-HDCim -Classe Win32_LogicalDisk -Filtro 'DriveType=3' | Select-Object DeviceID, VolumeName, FileSystem, Size, FreeSpace) }
        RiavvioInSospeso = { , @(Get-HDMotiviRiavvio) }
        Processi         = {
            , @(Get-Process | Group-Object -Property Name | ForEach-Object {
                    [pscustomobject]@{ Nome = $_.Name; Processi = $_.Count; MemoriaMB = [math]::Round((($_.Group | Measure-Object -Property WorkingSet64 -Sum).Sum) / 1MB) }
                } | Sort-Object -Property MemoriaMB -Descending | Select-Object -First 5)
        }
        Qfe              = { , @(Get-HDCim -Classe Win32_QuickFixEngineering | Select-Object HotFixID, InstalledOn) }
        Antivirus        = { , @(Get-HDCim -Classe AntiVirusProduct -Namespace 'root/SecurityCenter2' | Select-Object displayName, productState) }
        Defender         = {
            if (-not (Get-Command -Name Get-MpComputerStatus -ErrorAction SilentlyContinue)) { throw 'Modulo Defender non presente' }
            Get-MpComputerStatus -ErrorAction Stop | Select-Object AMServiceEnabled, AntivirusEnabled, RealTimeProtectionEnabled, AntivirusSignatureAge
        }
        Firewall         = { , @(Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop | Select-Object @{ n = 'Name'; e = { [string]$_.Name } }, @{ n = 'Enabled'; e = { [string]$_.Enabled } }) }
        BitLocker        = { , @(Get-HDStatoBitLocker -Admin $dati.Admin) }
        Rete             = { , @(Get-HDCim -Classe Win32_NetworkAdapterConfiguration -Filtro 'IPEnabled=True' | Select-Object Description, IPAddress, DefaultIPGateway, DNSServerSearchOrder, DHCPEnabled) }
        PingGateway      = { Test-HDPingGateway -Schede $dati.Rete }
        Dns              = {
            if ($SenzaInternet) { return }
            $indirizzi = [System.Net.Dns]::GetHostAddresses('www.microsoft.com')
            [pscustomobject]@{ Ok = (@($indirizzi).Count -gt 0) }
        }
        Internet         = { if ($SenzaInternet) { return }; Test-HDInternet }
        Eventi           = { , @(Get-HDEventi -Filtro @{ LogName = 'System'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } -Massimo 1000) }
        EventiDisco      = {
            $tutti = @()
            foreach ($fornitore in 'disk', 'Microsoft-Windows-Ntfs') {
                try { $tutti += @(Get-HDEventi -Filtro @{ LogName = 'System'; ProviderName = $fornitore; Level = 1, 2, 3; StartTime = (Get-Date).AddDays(-30) } -Massimo 200) }
                catch { Write-Verbose "Eventi di '$fornitore' non leggibili: $($_.Exception.Message)" }
            }
            , $tutti
        }
        DischiFisici     = { , @(Get-HDCim -Classe MSFT_PhysicalDisk -Namespace 'root/Microsoft/Windows/Storage' | Select-Object FriendlyName, HealthStatus, OperationalStatus, MediaType, BusType) }
        Servizi          = { , @(Get-HDCim -Classe Win32_Service -Filtro $filtroServizi | Select-Object Name, DisplayName, State, StartMode) }
        Stampanti        = { , @(Get-HDCim -Classe Win32_Printer | Select-Object Name, Default, WorkOffline, PrinterStatus, PortName) }
        CodaStampa       = { , @(Get-HDCim -Classe Win32_PrintJob | Select-Object Name, Document, JobStatus, Status, TimeSubmitted) }
        PuntiRipristino  = {
            if (-not $dati.Admin) { throw 'Serve avviare il toolkit come amministratore' }
            , @(Get-HDCim -Classe SystemRestore -Namespace 'root/default' | Select-Object CreationTime, Description)
        }
    }

    $i = 0
    foreach ($nome in @($passi.Keys)) {
        $i++
        Write-Progress -Activity 'Verifica postazione' -Status "Controllo: $nome" -PercentComplete ([int]($i * 100 / $passi.Count))
        try { $dati[$nome] = & $passi[$nome] }
        catch {
            $dati[$nome] = $null
            $dati.Errori[$nome] = $_.Exception.Message
        }
    }
    Write-Progress -Activity 'Verifica postazione' -Completed
    $dati
}

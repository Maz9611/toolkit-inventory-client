# Dati di prova comuni ai test (da richiamare con  . (Join-Path $PSScriptRoot 'DatiDiProva.ps1')  dentro BeforeAll)

# Dati grezzi completi (anche quelli per i fogli Dettagli e Software)
function New-GrezzoCompleto {
    param([string]$Nome = 'PC-AMMINISTRAZIONE', [string]$Serie = 'SER-001', [string]$Mac = '00:11:22:33:44:01', [object[]]$Programmi)
    if (-not $PSBoundParameters.ContainsKey('Programmi')) {
        $Programmi = @(
            [pscustomobject]@{ DisplayName = 'Mozilla Firefox (x64 it)'; DisplayVersion = '131.0'; Publisher = 'Mozilla'; InstallDate = '20241005' }
            [pscustomobject]@{ DisplayName = 'Adobe Acrobat Reader'; DisplayVersion = '24.003'; Publisher = 'Adobe'; InstallDate = '20240110' }
            [pscustomobject]@{ DisplayName = 'TeamViewer'; DisplayVersion = '15.50'; Publisher = 'TeamViewer' }
            [pscustomobject]@{ DisplayName = 'Microsoft 365 Apps for business - it-it'; DisplayVersion = '16.0.17126.20132'; Publisher = 'Microsoft Corporation' }
            [pscustomobject]@{ DisplayName = 'Microsoft Visual C++ 2015-2022 Redistributable (x64) - 14.38.33135'; DisplayVersion = '14.38'; Publisher = 'Microsoft Corporation' }
            [pscustomobject]@{ DisplayName = 'Componente di sistema'; SystemComponent = 1 }
        )
    }
    [pscustomobject]@{
        NomePC = $Nome; UnitaSistema = 'C:'
        ComputerSystem = [pscustomobject]@{ Manufacturer = 'Dell Inc.'; Model = 'OptiPlex 7090'; Domain = 'studio.local'; PartOfDomain = $true; UserName = 'STUDIO\giulia'; TotalPhysicalMemory = 16GB; PCSystemType = 1 }
        Prodotto = $null
        SistemaOperativo = [pscustomobject]@{ Caption = 'Microsoft Windows 10 Pro'; BuildNumber = '19045' }
        VersioneWindows = [pscustomobject]@{ DisplayVersion = '22H2'; CurrentBuild = '19045'; UBR = 5011 }
        Bios = [pscustomobject]@{ SerialNumber = $Serie; SMBIOSBIOSVersion = '1.15.0'; ReleaseDate = [datetime]'2022-03-12' }
        Chassis = @(3)
        Processori = @([pscustomobject]@{ Name = 'Intel(R) Core(TM) i7-11700 @ 2.50GHz' })
        Memoria = @([pscustomobject]@{ Capacity = 8GB }, [pscustomobject]@{ Capacity = 8GB })
        Volumi = @([pscustomobject]@{ DeviceID = 'C:'; Size = 476GB; FreeSpace = 120GB })
        ConfigRete = @([pscustomobject]@{ Description = 'Intel Ethernet'; IPAddress = @('192.168.1.21'); MACAddress = $Mac; DefaultIPGateway = @('192.168.1.1'); DHCPEnabled = $true; IPConnectionMetric = 25 })
        SchedeRete = @(); UltimoAccesso = $null; Profili = @(); Entra = $null
        Cynet = [pscustomobject]@{ Processi = @('CynetEPS'); Servizi = @(); Programmi = @(); File = @() }
        Avvio = 2
        SecureBoot = [pscustomobject]@{ Attivo = 0 }
        Tpm = [pscustomobject]@{ SpecVersion = '2.0, 0, 1.38'; IsEnabled_InitialValue = $true }
        TpmDispositivi = $null
        DischiFisici = @(
            [pscustomobject]@{ FriendlyName = 'Samsung SSD 980 500GB'; MediaType = 4; BusType = 17; Size = 500GB; HealthStatus = 0 }
            [pscustomobject]@{ FriendlyName = 'WDC WD10EZEX'; MediaType = 3; BusType = 11; Size = 1000GB; HealthStatus = 1 }
            [pscustomobject]@{ FriendlyName = 'Chiavetta USB'; MediaType = 0; BusType = 7; Size = 32GB; HealthStatus = 0 }
        )
        BitLockerAdmin = @([pscustomobject]@{ DriveLetter = 'C:'; ProtectionStatus = 1 })
        Smart = @(
            [pscustomobject]@{ Disco = 'Samsung SSD 980 500GB'; Temperatura = 38; Usura = 3; Ore = 2140; ErroriLettura = 0; ErroriScrittura = 0 }
            [pscustomobject]@{ Disco = 'WDC WD10EZEX'; Temperatura = 41; Usura = $null; Ore = 31200; ErroriLettura = 12; ErroriScrittura = 0 }
        )
        SmartPrevisione = @([pscustomobject]@{ InstanceName = 'x'; PredictFailure = $false })
        Monitor = @(
            [pscustomobject]@{ Produttore = 'DEL'; Nome = 'DELL P2419H'; Seriale = 'CFV9N93'; Anno = 2020 }
            [pscustomobject]@{ Produttore = 'SAM'; Nome = 'S24F350'; Seriale = '0'; Anno = 2018 }
        )
        Stampanti = @(
            [pscustomobject]@{ Name = 'HP LaserJet M404'; DriverName = 'HP Universal Printing PCL 6'; PortName = 'IP_192.168.1.50'; Default = $true }
            [pscustomobject]@{ Name = 'Microsoft Print to PDF'; DriverName = 'Microsoft Print To PDF'; PortName = 'PORTPROMPT:' }
        )
        Programmi = $Programmi
        OfficeC2R = [pscustomobject]@{ ProductReleaseIds = 'O365BusinessRetail'; VersionToReport = '16.0.17126.20132' }
        Licenza = @([pscustomobject]@{ Name = 'Windows(R), Professional edition'; Description = 'Windows(R) Operating System, OEM_DM channel'; LicenseStatus = 1 })
        Antivirus = @([pscustomobject]@{ displayName = 'Windows Defender'; productState = 397568 })
        DesktopRemoto = 1
        Batteria = [pscustomobject]@{ Presenti = 0; Progetto = $null; Piena = $null }
        Qfe = @(
            [pscustomobject]@{ HotFixID = 'KB5065000'; InstalledOn = '9/9/2026' }
            [pscustomobject]@{ HotFixID = 'KB5066835'; InstalledOn = '10/14/2026' }
            [pscustomobject]@{ HotFixID = 'KB5000001'; InstalledOn = '' }
        )
        Amministratori = @(
            [pscustomobject]@{ Percorso = "WinNT://STUDIO/$Nome/Administrator"; Classe = 'User'; Disattivato = $true }
            [pscustomobject]@{ Percorso = 'WinNT://STUDIO/giulia'; Classe = 'User'; Disattivato = $null }
            [pscustomobject]@{ Percorso = 'WinNT://STUDIO/Domain Admins'; Classe = 'Group'; Disattivato = $null }
            [pscustomobject]@{ Percorso = 'WinNT://S-1-12-1-1111111111-2222222222-3333333333-4444444444'; Classe = 'Group'; Disattivato = $null }
        )
        Raccolto = [datetime]'2026-11-02 10:00'
    }
}

# Fogli dell'Excel letti dal modulo, e righe di un foglio (senza l'intestazione)
function Get-Fogli([string]$Percorso) {
    InModuleScope HelpDeskToolkit -Parameters @{ P = $Percorso } { param($P) Read-HDXlsx -Percorso $P }
}
function Get-RigheFoglio($Fogli, [string]$Nome) {
    $f = $Fogli | Where-Object { $_.Nome -eq $Nome }
    foreach ($n in $f.Righe.Keys) { if ($n -gt 1) { , $f.Righe[$n] } }
}

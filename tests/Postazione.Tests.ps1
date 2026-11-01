BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    $script:Adesso = [datetime]'2026-11-01 09:00'

    # Dati di un PC sano: ogni test ne rompe un pezzo
    function New-DatiFinti {
        $adesso = [datetime]'2026-11-01 09:00'
        @{
            Errori           = @{}
            Admin            = $true
            SenzaInternet    = $false
            Raccolto         = $adesso
            NomePC           = 'PC-TEST'
            Utente           = 'UFFICIO\mario'
            FusoOrario       = '(UTC+01:00) Roma'
            OS               = [pscustomobject]@{ Caption = 'Microsoft Windows 11 Pro'; Version = '10.0.26100'; BuildNumber = '26100'; LastBootUpTime = $adesso.AddDays(-2); FreePhysicalMemory = 8388608; TotalVisibleMemorySize = 16777216 }
            CS               = [pscustomobject]@{ Name = 'PC-TEST'; Domain = 'ufficio.local'; PartOfDomain = $true; Manufacturer = 'Dell Inc.'; Model = 'OptiPlex 7090'; UserName = 'UFFICIO\mario' }
            Bios             = [pscustomobject]@{ SerialNumber = 'ABC1234' }
            VersioneWindows  = [pscustomobject]@{ DisplayVersion = '24H2'; CurrentBuild = '26100'; UBR = 2033 }
            Dischi           = @([pscustomobject]@{ DeviceID = 'C:'; VolumeName = 'OS'; FileSystem = 'NTFS'; Size = 500GB; FreeSpace = 200GB })
            RiavvioInSospeso = @()
            Processi         = @([pscustomobject]@{ Nome = 'chrome'; Processi = 25; MemoriaMB = 2100 })
            Qfe              = @([pscustomobject]@{ HotFixID = 'KB5066835'; InstalledOn = '10/20/2026' })
            Antivirus        = @([pscustomobject]@{ displayName = 'Windows Defender'; productState = 397568 })
            Defender         = [pscustomobject]@{ AMServiceEnabled = $true; AntivirusEnabled = $true; RealTimeProtectionEnabled = $true; AntivirusSignatureAge = 0 }
            Firewall         = @([pscustomobject]@{ Name = 'Domain'; Enabled = 'True' }, [pscustomobject]@{ Name = 'Private'; Enabled = 'True' }, [pscustomobject]@{ Name = 'Public'; Enabled = 'True' })
            BitLocker        = @([pscustomobject]@{ MountPoint = 'C:'; Protezione = 'Non attiva'; Fonte = '' })
            Rete             = @([pscustomobject]@{ Description = 'Intel Ethernet'; IPAddress = @('192.168.1.50', 'fe80::2'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('192.168.1.10'); DHCPEnabled = $true })
            PingGateway      = @{ '192.168.1.1' = $true }
            Dns              = [pscustomobject]@{ Ok = $true }
            Internet         = [pscustomobject]@{ Ok = $true; DataServerUtc = [datetime]'2026-11-01 08:00:00'; OraLocaleUtc = [datetime]'2026-11-01 08:00:20' }
            Eventi           = @()
            EventiDisco      = @()
            DischiFisici     = @([pscustomobject]@{ FriendlyName = 'Samsung SSD 980'; HealthStatus = 0; OperationalStatus = 2; MediaType = 4; BusType = 17 })
            Servizi          = @([pscustomobject]@{ Name = 'Spooler'; DisplayName = 'Spooler'; State = 'Running'; StartMode = 'Auto' }, [pscustomobject]@{ Name = 'W32Time'; DisplayName = 'Ora'; State = 'Stopped'; StartMode = 'Manual' })
            Stampanti        = @([pscustomobject]@{ Name = 'HP LaserJet Ufficio'; Default = $true; WorkOffline = $false }, [pscustomobject]@{ Name = 'Microsoft Print to PDF'; Default = $false; WorkOffline = $false })
            CodaStampa       = @()
            PuntiRipristino  = @([pscustomobject]@{ CreationTime = '20261025100000.000000+060'; Description = 'Windows Update' })
        }
    }

    function Invoke-Controllo([string]$Nome, [hashtable]$Dati) {
        InModuleScope HelpDeskToolkit -Parameters @{ N = $Nome; D = $Dati; A = $script:Adesso } {
            param($N, $D, $A)
            & $script:HDControlli[$N] -Dati $D -Soglie (Get-HDConfig).Soglie -Adesso $A
        }
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Un PC sano' {
    It 'passa tutti i controlli senza errori' {
        $e = @(InModuleScope HelpDeskToolkit -Parameters @{ D = (New-DatiFinti); A = $script:Adesso } { param($D, $A) Invoke-HDControlli -Dati $D -Adesso $A })
        $e.Count | Should -BeGreaterOrEqual 19
        @($e | Where-Object { $_.Stato -in 'ERRORE', 'CRITICO', 'ATTENZIONE' }) | Should -HaveCount 0
        ($e | Where-Object Controllo -eq 'BitLocker').Stato | Should -Be 'INFO'
    }
}

Describe 'Singoli controlli' {
    It 'Windows 10 -> ATTENZIONE (fuori supporto)' {
        $d = New-DatiFinti; $d.OS.Caption = 'Microsoft Windows 10 Pro'
        $r = Invoke-Controllo 'SistemaOperativo' $d
        $r.Stato | Should -Be 'ATTENZIONE'
        $r.Suggerimento | Should -Match 'Windows 11'
    }

    It 'dato mancante -> NON VERIFICATO con il motivo' {
        $d = New-DatiFinti; $d.OS = $null; $d.Errori.OS = 'Accesso negato'
        $r = Invoke-Controllo 'SistemaOperativo' $d
        $r.Stato | Should -Be 'NON VERIFICATO'
        $r.Dettaglio | Should -Be 'Accesso negato'
    }

    It 'acceso da 20 giorni -> ATTENZIONE' {
        $d = New-DatiFinti; $d.OS.LastBootUpTime = $script:Adesso.AddDays(-20)
        (Invoke-Controllo 'Uptime' $d).Stato | Should -Be 'ATTENZIONE'
    }

    It 'riavvio in sospeso: <Motivi> -> <Atteso>' -TestCases @(
        @{ Motivi = @('Windows Update'); Atteso = 'ATTENZIONE' }
        @{ Motivi = @('Operazioni su file in attesa'); Atteso = 'INFO' }
        @{ Motivi = @(); Atteso = 'OK' }
    ) {
        $d = New-DatiFinti; $d.RiavvioInSospeso = $Motivi
        (Invoke-Controllo 'Riavvio' $d).Stato | Should -Be $Atteso
    }

    It 'disco con 3 GB liberi -> CRITICO' {
        $d = New-DatiFinti; $d.Dischi += [pscustomobject]@{ DeviceID = 'D:'; Size = 100GB; FreeSpace = 3GB }
        $r = Invoke-Controllo 'Disco' $d
        $r.Stato | Should -Be 'CRITICO'
        $r.Dettaglio | Should -Match 'D: 3,0 GB liberi'
    }

    It 'RAM: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'uso 95%'; TotaleKB = 16777216; LiberaKB = 838861; Atteso = 'ATTENZIONE' }
        @{ Caso = 'solo 2 GB'; TotaleKB = 2097152; LiberaKB = 1048576; Atteso = 'ATTENZIONE' }
        @{ Caso = 'normale'; TotaleKB = 16777216; LiberaKB = 8388608; Atteso = 'OK' }
    ) {
        $d = New-DatiFinti; $d.OS.TotalVisibleMemorySize = $TotaleKB; $d.OS.FreePhysicalMemory = $LiberaKB
        (Invoke-Controllo 'Memoria' $d).Stato | Should -Be $Atteso
    }

    It 'aggiornamenti di <Giorni> giorni fa -> <Atteso>' -TestCases @(
        @{ Giorni = 10; Atteso = 'OK' }
        @{ Giorni = 60; Atteso = 'ATTENZIONE' }
        @{ Giorni = 120; Atteso = 'CRITICO' }
    ) {
        $d = New-DatiFinti
        $d.Qfe = @([pscustomobject]@{ HotFixID = 'KB1'; InstalledOn = $script:Adesso.AddDays(-$Giorni).ToString('M/d/yyyy', [cultureinfo]::InvariantCulture) })
        (Invoke-Controllo 'Aggiornamenti' $d).Stato | Should -Be $Atteso
    }

    It 'antivirus: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'nessuno attivo'; Prodotti = @([pscustomobject]@{ displayName = 'Windows Defender'; productState = 393472 }); RealTime = $false; Eta = 0; Atteso = 'CRITICO' }
        @{ Caso = 'terze parti non aggiornato'; Prodotti = @([pscustomobject]@{ displayName = 'Avast'; productState = 266256 }); RealTime = $false; Eta = 0; Atteso = 'ATTENZIONE' }
        @{ Caso = 'firme Defender vecchie'; Prodotti = @([pscustomobject]@{ displayName = 'Windows Defender'; productState = 397568 }); RealTime = $true; Eta = 12; Atteso = 'ATTENZIONE' }
    ) {
        $d = New-DatiFinti; $d.Antivirus = $Prodotti; $d.Defender.RealTimeProtectionEnabled = $RealTime; $d.Defender.AntivirusSignatureAge = $Eta
        (Invoke-Controllo 'Antivirus' $d).Stato | Should -Be $Atteso
    }

    It 'firewall spento sul profilo pubblico -> ATTENZIONE' {
        $d = New-DatiFinti; $d.Firewall[2].Enabled = 'False'
        $r = Invoke-Controllo 'Firewall' $d
        $r.Stato | Should -Be 'ATTENZIONE'
        $r.Dettaglio | Should -Match 'Pubblico: DISATTIVATO'
    }

    It 'BitLocker attivo ricorda la chiave di ripristino, sospeso -> ATTENZIONE' {
        $d = New-DatiFinti; $d.BitLocker[0].Protezione = 'Attiva'
        $r = Invoke-Controllo 'BitLocker' $d
        $r.Stato | Should -Be 'INFO'
        $r.Suggerimento | Should -Match 'chiave di ripristino'
        $d.BitLocker[0].Protezione = 'Sospesa'
        (Invoke-Controllo 'BitLocker' $d).Stato | Should -Be 'ATTENZIONE'
    }

    It 'rete: indirizzo 169.254 -> CRITICO' {
        $d = New-DatiFinti; $d.Rete[0].IPAddress = @('169.254.10.20'); $d.Rete[0].DefaultIPGateway = $null
        $r = @(Invoke-Controllo 'Rete' $d)
        ($r | Where-Object Controllo -eq 'Indirizzo IP').Stato | Should -Be 'CRITICO'
        ($r | Where-Object Controllo -eq 'Gateway').Stato | Should -Be 'ATTENZIONE'
    }

    It 'rete: DNS e Internet che non rispondono -> CRITICO con il motivo' {
        $d = New-DatiFinti; $d.Dns = $null; $d.Internet = $null; $d.Errori.Internet = 'Timeout'
        $r = @(Invoke-Controllo 'Rete' $d)
        ($r | Where-Object Controllo -eq 'DNS').Stato | Should -Be 'CRITICO'
        ($r | Where-Object Controllo -eq 'Internet').Dettaglio | Should -Match 'Timeout'
    }

    It 'rete: in modalità senza Internet salta DNS e Internet' {
        $d = New-DatiFinti; $d.SenzaInternet = $true; $d.Dns = $null; $d.Internet = $null
        $r = @(Invoke-Controllo 'Rete' $d)
        ($r | Where-Object Controllo -eq 'Internet').Stato | Should -Be 'INFO'
    }

    It 'orologio avanti di 10 minuti -> ATTENZIONE' {
        $d = New-DatiFinti; $d.Internet.OraLocaleUtc = $d.Internet.DataServerUtc.AddMinutes(10)
        $r = Invoke-Controllo 'Ora' $d
        $r.Stato | Should -Be 'ATTENZIONE'
        $r.Dettaglio | Should -Match 'avanti di 10 minuti'
    }

    It 'errori del disco nel registro eventi -> CRITICO con invito al backup' {
        $d = New-DatiFinti; $d.EventiDisco = @([pscustomobject]@{ ProviderName = 'disk'; Id = 7; Level = 2; TimeCreated = $script:Adesso.AddDays(-1) })
        $r = Invoke-Controllo 'Eventi' $d
        $r.Stato | Should -Be 'CRITICO'
        $r.Suggerimento | Should -Match 'BACKUP'
    }

    It 'spegnimenti imprevisti -> ATTENZIONE; altri errori generici -> INFO' {
        $d = New-DatiFinti; $d.Eventi = @([pscustomobject]@{ ProviderName = 'Microsoft-Windows-Kernel-Power'; Id = 41; Level = 1; TimeCreated = $script:Adesso })
        (Invoke-Controllo 'Eventi' $d).Stato | Should -Be 'ATTENZIONE'
        $d.Eventi = @([pscustomobject]@{ ProviderName = 'Service Control Manager'; Id = 7000; Level = 2; TimeCreated = $script:Adesso })
        (Invoke-Controllo 'Eventi' $d).Stato | Should -Be 'INFO'
    }

    It 'salute disco <Valore> -> <Atteso>' -TestCases @(
        @{ Valore = 1; Atteso = 'ATTENZIONE' }
        @{ Valore = 'Unhealthy'; Atteso = 'CRITICO' }
        @{ Valore = 'Healthy'; Atteso = 'OK' }
    ) {
        $d = New-DatiFinti; $d.DischiFisici[0].HealthStatus = $Valore
        (Invoke-Controllo 'DischiFisici' $d).Stato | Should -Be $Atteso
    }

    It 'spooler fermo -> ATTENZIONE; spooler disabilitato -> INFO' {
        $d = New-DatiFinti; $d.Servizi[0].State = 'Stopped'
        (Invoke-Controllo 'Servizi' $d).Stato | Should -Be 'ATTENZIONE'
        $d.Servizi[0].StartMode = 'Disabled'
        (Invoke-Controllo 'Servizi' $d).Stato | Should -Be 'INFO'
    }

    It 'stampanti: lavori in coda o stampante offline -> ATTENZIONE' {
        $d = New-DatiFinti; $d.CodaStampa = @([pscustomobject]@{ Name = 'HP LaserJet Ufficio, 12'; Document = 'fattura.pdf'; JobStatus = 'Errore'; Status = 'Error' })
        $r = Invoke-Controllo 'Stampanti' $d
        $r.Stato | Should -Be 'ATTENZIONE'
        $r.Dettaglio | Should -Match '1 lavoro in coda \(1 in errore\)'
        $r.Dettaglio | Should -Match '^2 stampanti \(1 fisica\)'
        $d = New-DatiFinti; $d.Stampanti[0].WorkOffline = $true
        (Invoke-Controllo 'Stampanti' $d).Dettaglio | Should -Match 'Offline: HP LaserJet Ufficio'
    }

    It 'punti di ripristino: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'non amministratore'; Admin = $false; Punti = @(); Atteso = 'NON VERIFICATO' }
        @{ Caso = 'nessun punto'; Admin = $true; Punti = @(); Atteso = 'ATTENZIONE' }
        @{ Caso = 'punto vecchio'; Admin = $true; Punti = @([pscustomobject]@{ CreationTime = '20260101100000.000000+060' }); Atteso = 'INFO' }
    ) {
        $d = New-DatiFinti; $d.Admin = $Admin; $d.PuntiRipristino = $Punti
        (Invoke-Controllo 'PuntoRipristino' $d).Stato | Should -Be $Atteso
    }
}

Describe 'Conversione di date WMI e HTTP' {
    It 'data WMI con fuso -> ora UTC corretta' {
        $d = InModuleScope HelpDeskToolkit { ConvertFrom-HDDataWmi '20261015083000.000000+060' }
        $d.ToUniversalTime() | Should -Be ([datetime]'2026-10-15 07:30:00')
    }

    It 'intestazione HTTP Date -> UTC' {
        $d = InModuleScope HelpDeskToolkit { ConvertFrom-HDDataHttp 'Sun, 01 Nov 2026 14:00:00 GMT' }
        $d.Kind | Should -Be 'Utc'
        $d | Should -Be ([datetime]'2026-11-01 14:00:00')
    }
}

Describe 'Esecuzione dei controlli' {
    It '-Escludi salta i controlli indicati' {
        $e = @(InModuleScope HelpDeskToolkit -Parameters @{ D = (New-DatiFinti) } { param($D) Invoke-HDControlli -Dati $D -Escludi 'Rete', 'Eventi' })
        @($e | Where-Object Area -eq 'Rete') | Should -HaveCount 0
        @($e | Where-Object Controllo -eq 'Registro eventi') | Should -HaveCount 0
    }

    It 'un controllo che va in errore non blocca gli altri' {
        Mock -ModuleName HelpDeskToolkit Test-HDCtrlUptime { throw 'boom' }
        $e = @(InModuleScope HelpDeskToolkit -Parameters @{ D = (New-DatiFinti) } { param($D) Invoke-HDControlli -Dati $D })
        $err = @($e | Where-Object Stato -eq 'ERRORE')
        $err | Should -HaveCount 1
        $err[0].Controllo | Should -Be 'Uptime'
        $err[0].Dettaglio | Should -Match 'boom'
        $e.Count | Should -BeGreaterOrEqual 19
    }

    It 'Test-Postazione salva report HTML e JSON e fa l''escape dei testi' {
        Mock -ModuleName HelpDeskToolkit Test-HDWindows { $true }
        Mock -ModuleName HelpDeskToolkit Get-HDDatiPostazione {
            $d = New-DatiFinti
            $d.Stampanti[0].Name = 'Stampante <b>rotta</b>'
            $d
        }
        $e = @(Test-Postazione -PassThru -Silenzioso)
        $e.Count | Should -BeGreaterOrEqual 19
        $html = @(Get-ChildItem -LiteralPath $script:CartellaReportTest -Filter 'Verifica_PC-TEST_*.html')
        $html | Should -HaveCount 1
        $contenuto = [System.IO.File]::ReadAllText($html[0].FullName)
        $contenuto | Should -Match 'Stampante &lt;b&gt;rotta&lt;/b&gt;'
        $contenuto | Should -Not -Match '<b>rotta'
        $json = [System.IO.Path]::ChangeExtension($html[0].FullName, '.json')
        $dati = Get-Content -LiteralPath $json -Raw -Encoding UTF8 | ConvertFrom-Json
        $dati.Postazione.NomePC | Should -Be 'PC-TEST'
        @($dati.Controlli).Count | Should -Be $e.Count
    }

    It 'Test-Postazione si rifiuta di girare fuori da Windows' {
        Mock -ModuleName HelpDeskToolkit Test-HDWindows { $false }
        { Test-Postazione -Silenzioso } | Should -Throw '*solo su Windows*'
    }
}

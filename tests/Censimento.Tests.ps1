BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')

    # Dati grezzi di un PC come li restituisce la raccolta su Windows
    function New-GrezzoPC {
        param([string]$Nome = 'PC-RECEPTION', [string]$Mac = '00:11:22:33:44:55', [string]$MacWifi = '66:77:88:99:AA:BB', [string]$Ip = '192.168.1.20', [string]$Serie = 'CZC1234XYZ', [switch]$FuoriRete, [switch]$SenzaCynet)
        $config = @([pscustomobject]@{ Description = 'Intel(R) Ethernet Connection I219-LM'; IPAddress = @($Ip, 'fe80::1'); MACAddress = $Mac; DefaultIPGateway = @('192.168.1.1'); DHCPEnabled = $true; IPConnectionMetric = 25 })
        if ($FuoriRete) { $config = @() }
        $cynet = [pscustomobject]@{ Processi = @('CynetEPS', 'CynetMS'); Servizi = @('CynetLauncher'); Programmi = @(); File = @() }
        if ($SenzaCynet) { $cynet = [pscustomobject]@{ Processi = $null; Servizi = @(); Programmi = $null; File = @() } }
        [pscustomobject]@{
            NomePC           = $Nome
            UnitaSistema     = 'C:'
            ComputerSystem   = [pscustomobject]@{ Manufacturer = 'Hewlett-Packard'; Model = 'HP ProDesk 400 G7 Microtower PC'; Domain = 'studio.local'; PartOfDomain = $true; Workgroup = $null; UserName = 'STUDIO\mario.rossi'; TotalPhysicalMemory = 17046798336; PCSystemType = 1 }
            Prodotto         = [pscustomobject]@{ Name = 'HP ProDesk 400 G7'; Version = '' }
            SistemaOperativo = [pscustomobject]@{ Caption = 'Microsoft Windows 11 Pro'; Version = '10.0.26100'; BuildNumber = '26100' }
            VersioneWindows  = [pscustomobject]@{ DisplayVersion = '24H2'; ReleaseId = '2009'; CurrentBuild = '26100'; UBR = 2033 }
            Bios             = [pscustomobject]@{ SerialNumber = $Serie }
            Chassis          = @(3)
            Processori       = @([pscustomobject]@{ Name = 'Intel(R) Core(TM) i5-10500 CPU @ 3.10GHz' })
            Memoria          = @([pscustomobject]@{ Capacity = 8GB }, [pscustomobject]@{ Capacity = 8GB })
            Volumi           = @([pscustomobject]@{ DeviceID = 'C:'; Size = 511GB; FreeSpace = 128GB }, [pscustomobject]@{ DeviceID = 'D:'; Size = 1TB; FreeSpace = 10GB })
            ConfigRete       = $config
            SchedeRete       = @([pscustomobject]@{ Name = 'Intel(R) Ethernet Connection I219-LM'; MACAddress = $Mac }, [pscustomobject]@{ Name = 'Intel(R) Wi-Fi 6 AX201'; MACAddress = $MacWifi }, [pscustomobject]@{ Name = 'Hyper-V Virtual Ethernet Adapter'; MACAddress = '00:15:5D:00:00:01' })
            UltimoAccesso    = 'STUDIO\mario.rossi'
            Profili          = @()
            Entra            = $null
            Cynet            = $cynet
            Raccolto         = [datetime]'2026-11-01 09:30'
        }
    }
    function ConvertTo-Record($Grezzo, [string[]]$Tecnici = @('admin', 'administrator')) {
        InModuleScope HelpDeskToolkit -Parameters @{ G = $Grezzo; T = $Tecnici } { param($G, $T) ConvertTo-HDRecordCensimento -Grezzo $G -AccountTecnici $T }
    }
    function Read-Tabella([string]$Percorso) {
        InModuleScope HelpDeskToolkit -Parameters @{ P = $Percorso } { param($P) Get-HDTabellaCensimento $P }
    }
    function Get-Riga($Tabella, [string]$Hostname) { $Tabella.Righe | Where-Object { $_.Valori['Hostname'] -eq $Hostname } }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Dati del PC -> riga del censimento' {
    It 'compila tutte le colonne automatiche' {
        $r = ConvertTo-Record (New-GrezzoPC)
        $r.Hostname | Should -Be 'PC-RECEPTION'
        $r.IndirizzoIP | Should -Be '192.168.1.20'
        $r.TipoIP | Should -Be 'DHCP'
        $r.MAC | Should -Be '00:11:22:33:44:55'
        $r.DeviceType | Should -Be 'PC CLIENT'
        $r.Manufacturer | Should -Be 'HP'
        $r.Model | Should -Be 'HP ProDesk 400 G7 Microtower PC'
        $r.NomeUtente | Should -Be 'mario.rossi'
        $r.Utilizzatore | Should -Be ''
        $r.Dominio | Should -Be 'studio.local'
        $r.SerialNumber | Should -Be 'CZC1234XYZ'
        $r.OS | Should -Be 'Windows 11 Pro'
        $r.OSVersion | Should -Be '24H2 (26100.2033)'
        $r.Memory | Should -Be '16 GB (2 x 8 GB)'
        $r.Processor | Should -Be 'Intel Core i5-10500 @ 3.10GHz'
        $r.FreeDiskPct | Should -Be '(C:) 128 GB 25% (D:) 10 GB 1%'
        $r.DiskSize | Should -Be '(C:) 511 GB (D:) 1024 GB'
        $r.DiscoMinimoPct | Should -Be 1
        $r.Cynet | Should -Be 'Sì'
        $r.CynetInEsecuzione | Should -BeTrue
        $r.TuttiMAC | Should -Contain '66:77:88:99:AA:BB'
        $r.TuttiMAC | Should -Not -Contain '00:15:5D:00:00:01'
    }

    It 'Lenovo: modello leggibile invece del codice macchina; portatile' {
        $g = New-GrezzoPC
        $g.ComputerSystem.Manufacturer = 'LENOVO'
        $g.ComputerSystem.Model = '20XW00GBIX'
        $g.Prodotto.Version = 'ThinkPad T14 Gen 2i'
        $g.Chassis = @(10)
        $r = ConvertTo-Record $g
        $r.Manufacturer | Should -Be 'Lenovo'
        $r.Model | Should -Be 'ThinkPad T14 Gen 2i'
        $r.DeviceType | Should -Be 'PC CLIENT'
    }

    It 'tipo dispositivo: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'portatile'; Chassis = @(10); Modello = 'X'; Os = 'Microsoft Windows 11 Pro'; Atteso = 'PC CLIENT' }
        @{ Caso = 'macchina virtuale con Windows 11'; Chassis = @(1); Modello = 'Virtual Machine'; Os = 'Microsoft Windows 11 Pro'; Atteso = 'PC CLIENT' }
        @{ Caso = 'Windows Server'; Chassis = @(3); Modello = 'PowerEdge T150'; Os = 'Microsoft Windows Server 2022 Standard'; Atteso = 'SERVER' }
        @{ Caso = 'server da rack'; Chassis = @(23); Modello = 'X'; Os = 'Microsoft Windows 10 Pro'; Atteso = 'SERVER' }
    ) {
        $g = New-GrezzoPC
        $g.Chassis = $Chassis
        $g.ComputerSystem.Model = $Modello
        $g.SistemaOperativo.Caption = $Os
        (ConvertTo-Record $g).DeviceType | Should -Be $Atteso
    }

    It 'PC fuori rete: niente IP, MAC della scheda cablata' {
        $r = ConvertTo-Record (New-GrezzoPC -FuoriRete)
        $r.IndirizzoIP | Should -Be ''
        $r.TipoIP | Should -Be ''
        $r.MAC | Should -Be '00:11:22:33:44:55'
    }

    It 'sceglie la scheda con gateway e ignora VPN e indirizzi 169.254' {
        $g = New-GrezzoPC
        $g.ConfigRete = @(
            [pscustomobject]@{ Description = 'TAP-Windows Adapter V9 (VPN)'; IPAddress = @('10.8.0.5'); MACAddress = '00:FF:00:FF:00:FF'; DefaultIPGateway = @('10.8.0.1'); DHCPEnabled = $true; IPConnectionMetric = 1 }
            [pscustomobject]@{ Description = 'Intel(R) Wi-Fi 6 AX201'; IPAddress = @('169.254.3.4'); MACAddress = '66:77:88:99:AA:BB'; DefaultIPGateway = $null; DHCPEnabled = $true; IPConnectionMetric = 50 }
            [pscustomobject]@{ Description = 'Intel(R) Ethernet Connection I219-LM'; IPAddress = @('192.168.1.20'); MACAddress = '00:11:22:33:44:55'; DefaultIPGateway = @('192.168.1.1'); DHCPEnabled = $false; IPConnectionMetric = 25 }
        )
        $r = ConvertTo-Record $g
        $r.IndirizzoIP | Should -Be '192.168.1.20'
        $r.TipoIP | Should -Be 'Statico'
    }

    It 'nome utente: salta gli account dei tecnici' {
        $g = New-GrezzoPC
        $g.ComputerSystem.UserName = 'PC-RECEPTION\Admin'
        $g.UltimoAccesso = 'PC-RECEPTION\admin'
        $g.Profili = @(
            [pscustomobject]@{ LocalPath = 'C:\Users\admin'; LastUseTime = [datetime]'2026-11-01'; Special = $false }
            [pscustomobject]@{ LocalPath = 'C:\Users\giulia'; LastUseTime = [datetime]'2026-10-30'; Special = $false }
            [pscustomobject]@{ LocalPath = 'C:\Users\luca'; LastUseTime = [datetime]'2026-09-01'; Special = $false }
            [pscustomobject]@{ LocalPath = 'C:\Windows\ServiceProfiles\LocalService'; LastUseTime = [datetime]'2026-11-01'; Special = $true }
        )
        $r = ConvertTo-Record $g
        $r.NomeUtente | Should -Be 'giulia'
        $r.FonteNomeUtente | Should -Be 'profilo usato più di recente'
    }

    It 'dominio: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'gruppo di lavoro'; InDominio = $false; Entra = $null; Atteso = 'No' }
        @{ Caso = 'Entra ID'; InDominio = $false; Entra = 'Studio Rossi Srl'; Atteso = 'Entra ID (Studio Rossi Srl)' }
        @{ Caso = 'dominio'; InDominio = $true; Entra = 'X'; Atteso = 'studio.local' }
    ) {
        $g = New-GrezzoPC
        $g.ComputerSystem.PartOfDomain = $InDominio
        $g.ComputerSystem.Workgroup = 'UFFICIO'
        $g.Entra = $Entra
        (ConvertTo-Record $g).Dominio | Should -Be $Atteso
    }

    It 'Cynet: installato ma fermo = Sì (con avviso), assente = No' {
        $g = New-GrezzoPC
        $g.Cynet = [pscustomobject]@{ Processi = @(); Servizi = @('CynetLauncher'); Programmi = @('Cynet EPS'); File = @() }
        $r = ConvertTo-Record $g
        $r.Cynet | Should -Be 'Sì'
        $r.CynetInEsecuzione | Should -BeFalse
        (ConvertTo-Record (New-GrezzoPC -SenzaCynet)).Cynet | Should -Be 'No'
    }

    It 'memoria senza banchi leggibili (macchine virtuali) dal totale' {
        $g = New-GrezzoPC
        $g.Memoria = $null
        $g.ComputerSystem.TotalPhysicalMemory = 4294430720
        (ConvertTo-Record $g).Memory | Should -Be '4 GB'
    }

    It 'memoria: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'portatile DDR4'; Moduli = @(@{ Capacity = 8GB; FormFactor = 12; SMBIOSMemoryType = 26 }, @{ Capacity = 8GB; FormFactor = 12; SMBIOSMemoryType = 26 }); Atteso = '16 GB DDR4 SODIMM (2 x 8 GB)' }
        @{ Caso = 'desktop DDR5'; Moduli = @(@{ Capacity = 32GB; FormFactor = 8; SMBIOSMemoryType = 34 }); Atteso = '32 GB DDR5 DIMM' }
        @{ Caso = 'banchi diversi'; Moduli = @(@{ Capacity = 8GB; FormFactor = 8; SMBIOSMemoryType = 24 }, @{ Capacity = 4GB; FormFactor = 8; SMBIOSMemoryType = 24 }); Atteso = '12 GB DDR3 DIMM (1 x 8 GB + 1 x 4 GB)' }
        @{ Caso = 'Windows vecchio (solo MemoryType)'; Moduli = @(@{ Capacity = 4GB; FormFactor = 8; SMBIOSMemoryType = 0; MemoryType = 24 }); Atteso = '4 GB DDR3 DIMM' }
        @{ Caso = 'memoria saldata LPDDR5'; Moduli = @(@{ Capacity = 16GB; FormFactor = 0; SMBIOSMemoryType = 35 }); Atteso = '16 GB LPDDR5' }
    ) {
        $g = New-GrezzoPC
        $g.Memoria = @($Moduli | ForEach-Object { [pscustomobject]$_ })
        (ConvertTo-Record $g).Memory | Should -Be $Atteso
    }

    It 'numero di serie: <Caso> -> <Atteso>' -TestCases @(
        @{ Caso = 'BIOS'; Bios = 'CZC1234XYZ'; Telaio = 'ALTRO'; Prodotto = ''; Scheda = ''; Atteso = 'CZC1234XYZ' }
        @{ Caso = 'BIOS segnaposto, c''è nel telaio'; Bios = 'To Be Filled By O.E.M.'; Telaio = 'TEL123'; Prodotto = ''; Scheda = ''; Atteso = 'TEL123' }
        @{ Caso = 'solo nel prodotto'; Bios = 'Default string'; Telaio = 'Default string'; Prodotto = 'PRD456'; Scheda = ''; Atteso = 'PRD456' }
        @{ Caso = 'PC assemblato: scheda madre'; Bios = 'System Serial Number'; Telaio = 'Chassis Serial Number'; Prodotto = 'System Serial Number'; Scheda = '190912345678'; Atteso = 'MB 190912345678' }
        @{ Caso = 'nessuno'; Bios = 'To Be Filled By O.E.M.'; Telaio = ''; Prodotto = ''; Scheda = 'Default string'; Atteso = '' }
    ) {
        $g = New-GrezzoPC
        $g.Bios = [pscustomobject]@{ SerialNumber = $Bios }
        $g | Add-Member -NotePropertyName Telaio -NotePropertyValue ([pscustomobject]@{ SerialNumber = $Telaio })
        $g | Add-Member -NotePropertyName SchedaMadre -NotePropertyValue ([pscustomobject]@{ SerialNumber = $Scheda })
        $g.Prodotto = [pscustomobject]@{ Version = ''; IdentifyingNumber = $Prodotto }
        $r = ConvertTo-Record $g
        $r.SerialNumber | Should -Be $Atteso
        $r.SerialeBios | Should -Be $Bios
    }

    It 'valori di Device Type e Dominio presi da config.psd1' {
        $g = New-GrezzoPC
        $g.ComputerSystem.PartOfDomain = $false
        $r = InModuleScope HelpDeskToolkit -Parameters @{ G = $g } { param($G) ConvertTo-HDRecordCensimento -Grezzo $G -TipoClient 'WORKSTATION' -SenzaDominio 'NESSUNO' }
        $r.DeviceType | Should -Be 'WORKSTATION'
        $r.Dominio | Should -Be 'NESSUNO'
    }
}

Describe 'Intestazioni dell''Excel' {
    It 'riconosce le colonne anche con nomi diversi o con aggiunte' {
        $m = InModuleScope HelpDeskToolkit {
            Find-HDColonneCensimento @{ 1 = 'Nome PC'; 2 = 'Indirizzo IP'; 3 = 'Tipo IP'; 4 = 'Tipo'; 5 = 'Produttore'; 6 = 'Numero di serie'; 7 = 'Cynet (sì o no)'; 8 = 'Free Disk Space %'; 9 = 'Operating System'; 10 = 'Operating System Version'; 11 = 'Colonna mia' }
        }
        $m.Hostname | Should -Be 1
        $m.TipoIP | Should -Be 3
        $m.DeviceType | Should -Be 4
        $m.Manufacturer | Should -Be 5
        $m.SerialNumber | Should -Be 6
        $m.Cynet | Should -Be 7
        $m.FreeDiskPct | Should -Be 8
        $m.OS | Should -Be 9
        $m.OSVersion | Should -Be 10
        $m.Values | Should -Not -Contain 11
    }
}

Describe 'Abbinamento e unione dei dati' {
    BeforeAll {
        $script:Righe = @(
            [pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-A'; MAC = '00-11-22-33-44-55'; IndirizzoIP = '192.168.1.10' } }
            [pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'PC-B'; SerialNumber = 'SER-B' } }
            [pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'PC-C' } }
            [pscustomobject]@{ Numero = 5; Valori = @{ IndirizzoIP = '192.168.1.99' } }
        )
    }

    It 'cerca per <Caso>' -TestCases @(
        @{ Caso = 'MAC (anche scritto diverso)'; Mac = @('00:11:22:33:44:55'); Serie = ''; Nome = ''; Ip = ''; Riga = 2 }
        @{ Caso = 'numero di serie'; Mac = @('AA:AA:AA:AA:AA:AA'); Serie = 'ser-b'; Nome = ''; Ip = ''; Riga = 3 }
        @{ Caso = 'hostname (anche FQDN)'; Mac = @(); Serie = ''; Nome = 'pc-c.studio.local'; Ip = ''; Riga = 4 }
        @{ Caso = 'IP solo su righe senza nome e MAC'; Mac = @(); Serie = ''; Nome = ''; Ip = '192.168.1.99'; Riga = 5 }
    ) {
        $t = InModuleScope HelpDeskToolkit -Parameters @{ R = $script:Righe; M = $Mac; S = $Serie; N = $Nome; I = $Ip } {
            param($R, $M, $S, $N, $I) Find-HDRigaCensimento -Righe $R -Mac $M -Seriale $S -Hostname $N -IP $I -ConsentiIP
        }
        $t.Riga.Numero | Should -Be $Riga
    }

    It 'docking station: lo stesso MAC su un altro portatile non sovrascrive la riga sbagliata' {
        $righe = @(
            [pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'LAPTOP-A'; MAC = 'D0:C0:00:00:00:01'; SerialNumber = 'SN-A' } }
            [pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'LAPTOP-C'; MAC = 'D0:C0:00:00:00:02' } }
            [pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'LAPTOP-B' } }
        )
        InModuleScope HelpDeskToolkit -Parameters @{ R = $righe } {
            param($R)
            # PC con un altro numero di serie: la riga di A non va toccata, si trova la sua per nome
            (Find-HDRigaCensimento -Righe $R -Mac @('D0:C0:00:00:00:01') -Seriale 'SN-B' -Hostname 'LAPTOP-B').Riga.Numero | Should -Be 4
            # riga venuta solo dalla scansione (senza seriale) con un altro nome: nessun abbinamento per MAC
            Find-HDRigaCensimento -Righe $R -Mac @('D0:C0:00:00:00:02') -Seriale 'SN-X' -Hostname 'LAPTOP-X' | Should -BeNullOrEmpty
            # stesso PC: il seriale conferma anche se nel frattempo è stato rinominato
            (Find-HDRigaCensimento -Righe $R -Mac @('D0:C0:00:00:00:01') -Seriale 'SN-A' -Hostname 'LAPTOP-NUOVO').Riga.Numero | Should -Be 2
        }
    }

    It 'nessun abbinamento -> riga nuova' {
        InModuleScope HelpDeskToolkit -Parameters @{ R = $script:Righe } { param($R) Find-HDRigaCensimento -Righe $R -Mac @('DE:AD:BE:EF:00:00') -Seriale 'X' -Hostname 'PC-Z' } | Should -BeNullOrEmpty
    }

    It 'dati del PC: niente valori vuoti al posto di quelli esistenti, colonne manuali solo se compilate' {
        $rec = ConvertTo-Record (New-GrezzoPC -FuoriRete)
        $diff = InModuleScope HelpDeskToolkit -Parameters @{ Rec = $rec } {
            param($Rec)
            Merge-HDValoriPC -Esistenti @{ IndirizzoIP = '192.168.1.20'; TipoIP = 'DHCP'; Utilizzatore = 'Mario'; Note = 'vecchia' } -Record $Rec -Manuali @{ Note = 'nuova'; Locazione = '' }
        }
        $diff.ContainsKey('IndirizzoIP') | Should -BeFalse
        $diff.ContainsKey('TipoIP') | Should -BeFalse
        $diff.ContainsKey('Utilizzatore') | Should -BeFalse
        $diff.ContainsKey('Locazione') | Should -BeFalse
        $diff.Note | Should -Be 'nuova'
        $diff.Cynet | Should -Be 'Sì'
    }

    It 'dati del PC: se la riga ha già un MAC di questo PC (es. il Wi-Fi) lo lascia' {
        $rec = ConvertTo-Record (New-GrezzoPC)
        $diff = InModuleScope HelpDeskToolkit -Parameters @{ Rec = $rec } { param($Rec) Merge-HDValoriPC -Esistenti @{ MAC = '66-77-88-99-AA-BB' } -Record $Rec }
        $diff.ContainsKey('MAC') | Should -BeFalse
    }

    It 'scansione: IP sempre aggiornato, il resto solo se manca' {
        $d = [pscustomobject]@{ Hostname = 'PC-A'; IP = '192.168.1.50'; MAC = '00:11:22:33:44:55'; Produttore = 'Intel Corporate'; Utente = ''; Gruppo = '' }
        $diff = InModuleScope HelpDeskToolkit -Parameters @{ D = $d } { param($D) Merge-HDValoriScansione -Esistenti @{ Hostname = 'PC-A'; IndirizzoIP = '192.168.1.10'; Manufacturer = 'HP' } -Dispositivo $D }
        $diff.IndirizzoIP | Should -Be '192.168.1.50'
        $diff.MAC | Should -Be '00:11:22:33:44:55'
        $diff.ContainsKey('Manufacturer') | Should -BeFalse
        $diff.ContainsKey('Hostname') | Should -BeFalse
        $d2 = [pscustomobject]@{ Hostname = 'PC-Z'; IP = '192.168.1.60'; MAC = ''; Produttore = ''; Utente = ''; Gruppo = 'WORKGROUP' }
        $diff2 = InModuleScope HelpDeskToolkit -Parameters @{ D = $d2 } { param($D) Merge-HDValoriScansione -Esistenti @{} -Dispositivo $D }
        $diff2.ContainsKey('Dominio') | Should -BeFalse
    }
}

Describe 'Dal file della scansione all''Excel completato sul PC' {
    BeforeAll {
        $script:Scansione = Join-Path $TestDrive 'scansione.csv'
        Set-Content -LiteralPath $script:Scansione -Encoding UTF8 -Value @(
            'Status,Name,IP,Manufacturer,MAC address,Comments'
            'On,pc-reception.studio.local,192.168.1.20,Intel Corporate,66:77:88:99:AA:BB,'
            'On,NPI3D4E5F,192.168.1.50,Hewlett Packard,AA:BB:CC:DD:EE:01,'
            'Off,,192.168.1.77,,,'
        )
        $script:FileCensimento = Join-Path $script:CartellaCensimentoTest 'Censimento_Studio-Rossi.xlsx'
    }

    It 'New-Censimento crea l''Excel con le 20 colonne nell''ordine giusto' {
        $f = New-Censimento -Cliente 'Prova Colonne'
        $f.Name | Should -Be 'Censimento_Prova-Colonne.xlsx'
        $fogli = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $f.FullName } { param($P) Read-HDXlsx $P })
        $titoli = @($fogli[0].Righe[1].GetEnumerator() | Sort-Object Key | ForEach-Object { $_.Value })
        $titoli | Should -Be @('Hostname', 'Indirizzo IP', 'Tipo IP', 'MAC Address', 'Device Type', 'Manufacturer', 'Model', 'Nome Utente', 'Utilizzatore', 'Note', 'Dominio', 'Serial Number', 'Operating System', 'Operating System Version', 'Memory', 'Processor Type', 'Free Disk Space %', 'Logical Disk Size', 'Cynet', 'Locazione')
        Remove-Item -LiteralPath $f.FullName
    }

    It 'l''import della scansione crea il censimento del cliente e aggiunge i dispositivi attivi' {
        $r = Import-ScansioneRete -Percorso $script:Scansione -Cliente 'Studio Rossi' -NonChiedere
        $r.Censimento | Should -Be $script:FileCensimento
        $r.Aggiunti | Should -Be 2
        $r.Scartati | Should -Be 1
        $t = Read-Tabella $script:FileCensimento
        $t.Righe | Should -HaveCount 2
        (Get-Riga $t 'PC-RECEPTION').Valori.MAC | Should -Be '66:77:88:99:AA:BB'
        (Get-Riga $t 'NPI3D4E5F').Valori.Manufacturer | Should -Be 'Hewlett Packard'
    }

    It 'il censimento sul PC ritrova la riga (MAC del Wi-Fi) e la completa senza toccare le colonne manuali' {
        $t = Read-Tabella $script:FileCensimento
        $n = (Get-Riga $t 'PC-RECEPTION').Numero
        InModuleScope HelpDeskToolkit -Parameters @{ P = $script:FileCensimento; T = $t; N = $n } {
            param($P, $T, $N) [void](Update-HDXlsx -Percorso $P -ParteFoglio $T.ParteFoglio -IndiceFoglio 0 -Modifiche @(@{ Riga = $N; Celle = @{ [int]$T.Colonne.Note = 'scritta a mano' } }))
        }
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoPC }
        $e = Add-PCCensimento -NonChiedere -Utilizzatore 'Mario Rossi' -Locazione 'Reception' 6>$null
        $e.Scritto | Should -BeTrue
        $e.Esiti[0].Azione | Should -Be 'Aggiornato'
        $e.Esiti[0].Motivo | Should -Match 'MAC 66:77:88:99:AA:BB'
        $t = Read-Tabella $script:FileCensimento
        $t.Righe | Should -HaveCount 2
        $v = (Get-Riga $t 'PC-RECEPTION').Valori
        $v.MAC | Should -Be '66:77:88:99:AA:BB'
        $v.SerialNumber | Should -Be 'CZC1234XYZ'
        $v.Memory | Should -Be '16 GB (2 x 8 GB)'
        $v.FreeDiskPct | Should -Be '(C:) 128 GB 25% (D:) 10 GB 1%'
        $v.Cynet | Should -Be 'Sì'
        $v.Utilizzatore | Should -Be 'Mario Rossi'
        $v.Locazione | Should -Be 'Reception'
        $v.Note | Should -Be 'scritta a mano'
        @(Get-ChildItem -LiteralPath (Join-Path $script:CartellaCensimentoTest 'backup') -Filter '*.xlsx').Count | Should -BeGreaterThan 0
        $rilevazioni = Join-Path (Join-Path $script:CartellaCensimentoTest 'rilevazioni') 'Censimento_Studio-Rossi'
        @(Get-ChildItem -LiteralPath $rilevazioni -Filter '*.json') | Should -HaveCount 1
        @(Get-ChildItem -LiteralPath (Join-Path $rilevazioni 'da-inserire') -Filter '*.json') | Should -HaveCount 0
    }

    It 'un PC che non era in rete viene aggiunto in fondo' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoPC -Nome 'PC-MAGAZZINO' -Mac 'DE:AD:BE:EF:00:01' -MacWifi 'DE:AD:BE:EF:00:02' -Serie 'MXL999' -FuoriRete -SenzaCynet }
        $e = Add-PCCensimento -NonChiedere 6>$null
        $e.Esiti[0].Azione | Should -Be 'Aggiunto'
        $t = Read-Tabella $script:FileCensimento
        $t.Righe | Should -HaveCount 3
        $v = (Get-Riga $t 'PC-MAGAZZINO').Valori
        $v.MAC | Should -Be 'DE:AD:BE:EF:00:01'
        $v.Cynet | Should -Be 'No'
        $v.ContainsKey('IndirizzoIP') | Should -BeFalse
    }

    It 'una nuova scansione aggiorna gli IP senza creare doppioni' {
        Set-Content -LiteralPath $script:Scansione -Encoding UTF8 -Value @('Name,IP,MAC address', 'PC-RECEPTION,192.168.1.120,00:11:22:33:44:55', 'NPI3D4E5F,192.168.1.50,AA:BB:CC:DD:EE:01')
        $r = Import-ScansioneRete -Percorso $script:Scansione -NonChiedere -Censimento $script:FileCensimento
        $r.Aggiunti | Should -Be 0
        $r.Aggiornati | Should -Be 1
        $r.Invariati | Should -Be 1
        $t = Read-Tabella $script:FileCensimento
        $t.Righe | Should -HaveCount 3
        (Get-Riga $t 'PC-RECEPTION').Valori.IndirizzoIP | Should -Be '192.168.1.120'
        (Get-Riga $t 'PC-RECEPTION').Valori.Utilizzatore | Should -Be 'Mario Rossi'
    }

    It 'se l''Excel non si può scrivere i dati restano da parte e Sync-Censimento li inserisce dopo' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoPC -Nome 'PC-UFFICIO2' -Mac 'AB:CD:EF:00:00:02' -MacWifi 'AB:CD:EF:00:00:03' -Ip '192.168.1.21' -Serie 'PF2ABCDE' }
        Mock -ModuleName HelpDeskToolkit Save-HDCensimento { throw 'file aperto in Excel' }
        $e = Add-PCCensimento -NonChiedere -Locazione 'Primo piano' 6>$null
        $e.Scritto | Should -BeFalse
        $e.Errore | Should -Match 'aperto'
        $attesa = Join-Path (Join-Path (Join-Path $script:CartellaCensimentoTest 'rilevazioni') 'Censimento_Studio-Rossi') 'da-inserire'
        @(Get-ChildItem -LiteralPath $attesa -Filter '*.json') | Should -HaveCount 1
        (Read-Tabella $script:FileCensimento).Righe | Should -HaveCount 3
    }

    It 'Sync-Censimento inserisce le rilevazioni rimaste in sospeso' {
        $s = Sync-Censimento -Censimento $script:FileCensimento
        $s.Inserite | Should -Be 1
        $t = Read-Tabella $script:FileCensimento
        $t.Righe | Should -HaveCount 4
        (Get-Riga $t 'PC-UFFICIO2').Valori.Locazione | Should -Be 'Primo piano'
    }

    It 'con più censimenti sceglie quello del cliente o il più recente' {
        $altro = New-Censimento -Cliente 'Altro Cliente'
        (Get-Item -LiteralPath $altro.FullName).LastWriteTime = (Get-Date).AddDays(-3)
        InModuleScope HelpDeskToolkit { Resolve-HDFileCensimento -Cliente 'Altro Cliente' -NonChiedere } | Should -Be $altro.FullName
        InModuleScope HelpDeskToolkit { Resolve-HDFileCensimento -NonChiedere } | Should -Be $script:FileCensimento
    }

    It 'usa il Modello.xlsx dell''utente se c''è' {
        $modello = Join-Path $script:CartellaCensimentoTest 'Modello.xlsx'
        Copy-Item -LiteralPath $script:FileCensimento -Destination $modello
        $f = New-Censimento -Cliente 'Da Modello'
        (Read-Tabella $f.FullName).Righe.Count | Should -Be 4
        Remove-Item -LiteralPath $modello, $f.FullName
    }

    It 'Get-InventarioPC avvisa se la raccolta non riesce' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { throw 'Accesso negato' }
        $r = Get-InventarioPC -ComputerName 'PC-REMOTO' -WarningVariable avvisi -WarningAction SilentlyContinue
        $r | Should -BeNullOrEmpty
        "$avvisi" | Should -Match 'PC-REMOTO.*Accesso negato'
    }
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')

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
            Raccolto = [datetime]'2026-11-02 10:00'
        }
    }
    function Get-Fogli([string]$Percorso) {
        InModuleScope HelpDeskToolkit -Parameters @{ P = $Percorso } { param($P) Read-HDXlsx -Percorso $P }
    }
    function Get-RigheFoglio($Fogli, [string]$Nome) {
        $f = $Fogli | Where-Object { $_.Nome -eq $Nome }
        foreach ($n in $f.Righe.Keys) { if ($n -gt 1) { , $f.Righe[$n] } }
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Pronto per Windows 11' {
    It 'processore <Nome> -> "<Atteso>"' -TestCases @(
        @{ Nome = 'Intel(R) Core(TM) i5-8250U CPU @ 1.60GHz'; Atteso = 'Sì' }
        @{ Nome = 'Intel(R) Core(TM) i7-7700HQ CPU @ 2.80GHz'; Atteso = 'No' }
        @{ Nome = 'Intel(R) Core(TM) i5-10500 CPU @ 3.10GHz'; Atteso = 'Sì' }
        @{ Nome = '11th Gen Intel(R) Core(TM) i7-1165G7 @ 2.80GHz'; Atteso = 'Sì' }
        @{ Nome = 'Intel(R) Core(TM) i3-4130 CPU @ 3.40GHz'; Atteso = 'No' }
        @{ Nome = 'Intel(R) Core(TM) i5 CPU 750 @ 2.67GHz'; Atteso = '' }
        @{ Nome = 'Intel(R) Core(TM) Ultra 7 155H'; Atteso = 'Sì' }
        @{ Nome = 'Intel(R) Core(TM) 5 120U'; Atteso = 'Sì' }
        @{ Nome = 'Intel(R) N100'; Atteso = 'Sì' }
        @{ Nome = 'AMD Ryzen 5 3500U with Radeon Vega Mobile Gfx'; Atteso = 'Sì' }
        @{ Nome = 'AMD Ryzen 7 PRO 4750U with Radeon Graphics'; Atteso = 'Sì' }
        @{ Nome = 'AMD Ryzen 7 1700 Eight-Core Processor'; Atteso = 'No' }
        @{ Nome = 'AMD Ryzen 5 2600 Six-Core Processor'; Atteso = 'Sì' }
        @{ Nome = 'AMD Ryzen 5 2500U with Radeon Vega Mobile Gfx'; Atteso = '' }
        @{ Nome = 'Intel(R) Core(TM)2 Duo CPU E8400 @ 3.00GHz'; Atteso = 'No' }
        @{ Nome = 'Intel(R) Celeron(R) N4020 CPU @ 1.10GHz'; Atteso = '' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ N = $Nome } { param($N) Test-HDCpuWindows11 $N } | Should -BeExactly $Atteso
    }

    It 'mette insieme TPM, avvio, Secure Boot, processore, RAM e disco' {
        InModuleScope HelpDeskToolkit {
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 11 Pro' | Should -Be 'Già Windows 11'
            Get-HDProntoWindows11 -SistemaOperativo 'Windows Server 2019 Standard' -Server $true | Should -Be ''
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 10 Pro' -Tpm '2.0' -Avvio 'UEFI' -SecureBoot 'Attivo' -Cpu 'Sì' -RamGB 16 -DiscoGB 476 | Should -Be 'Sì'
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 10 Pro' -Tpm '2.0' -Avvio 'UEFI' -SecureBoot 'Disattivato' -Cpu 'Sì' -RamGB 8 -DiscoGB 237 | Should -BeLike 'Sì (prima attiva Secure Boot*'
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 10 Pro' -Tpm '1.2' -Avvio 'BIOS' -SecureBoot 'Non supportato' -Cpu 'No' -RamGB 4 -DiscoGB 237 | Should -Be 'No: TPM 1.2 (serve 2.0), avvio BIOS legacy (serve UEFI), processore non supportato'
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 10 Pro' -Tpm '' -Avvio 'UEFI' -SecureBoot 'Attivo' -Cpu 'Sì' -RamGB 8 -DiscoGB 237 | Should -BeLike 'Da verificare: versione del TPM*'
            Get-HDProntoWindows11 -SistemaOperativo 'Windows 10 Pro' -Tpm '2.0' -Avvio 'UEFI' -SecureBoot 'Attivo' -Cpu 'Sì' -RamGB 2 -DiscoGB 32 | Should -Be 'No: RAM sotto 4 GB, disco sotto 64 GB'
        }
    }

    It 'legge il TPM da WMI (amministratore) o dai dispositivi (utente normale)' {
        InModuleScope HelpDeskToolkit {
            ConvertFrom-HDTpm -Wmi ([pscustomobject]@{ SpecVersion = '2.0, 0, 1.38'; IsEnabled_InitialValue = $true }) | Should -Be '2.0'
            ConvertFrom-HDTpm -Wmi ([pscustomobject]@{ SpecVersion = '1.2, 2, 3'; IsEnabled_InitialValue = $true }) | Should -Be '1.2'
            ConvertFrom-HDTpm -Wmi ([pscustomobject]@{ SpecVersion = '2.0, 0, 1.38'; IsEnabled_InitialValue = $false }) | Should -Be '2.0 (disattivato)'
            ConvertFrom-HDTpm -Dispositivi ([pscustomobject]@{ Nomi = @('Trusted Platform Module 2.0') }) | Should -Be '2.0'
            ConvertFrom-HDTpm -Dispositivi ([pscustomobject]@{ Nomi = @() }) | Should -Be 'Assente'
            ConvertFrom-HDTpm | Should -Be ''
        }
    }
}

Describe 'Hardware, software e licenze -> testi del foglio Dettagli' {
    BeforeAll {
        $script:Det = (InModuleScope HelpDeskToolkit -Parameters @{ G = (New-GrezzoCompleto) } { param($G) ConvertTo-HDDettagliPC -Grezzo $G -SistemaOperativo 'Windows 10 Pro' -Processore $G.Processori[0].Name })
    }

    It 'pronto per Windows 11, TPM, Secure Boot, avvio' {
        $script:Det.Dettagli.Windows11 | Should -BeLike 'Sì (prima attiva Secure Boot*'
        $script:Det.Dettagli.TPM | Should -Be '2.0'
        $script:Det.Dettagli.SecureBoot | Should -Be 'Disattivato'
        $script:Det.Dettagli.Avvio | Should -Be 'UEFI'
    }

    It 'dischi fisici senza le chiavette USB, con tipo e salute' {
        $script:Det.Dettagli.DischiFisici | Should -Be 'Samsung SSD 980 500GB (SSD NVMe, 500 GB, OK); WDC WD10EZEX (HDD SATA, 1000 GB, ATTENZIONE)'
    }

    It 'SMART dei dischi: usura, temperatura, ore ed errori' {
        $script:Det.Dettagli.Smart | Should -Be 'Samsung SSD 980 500GB: usura 3%, 38 °C, 2.140 ore, nessun errore; WDC WD10EZEX: 41 °C, 31.200 ore, ATTENZIONE 12 errori non corretti'
        InModuleScope HelpDeskToolkit {
            Format-HDSmart -Contatori @([pscustomobject]@{ Disco = 'SSD'; Usura = 1 }) -Previsione @([pscustomobject]@{ PredictFailure = $true }) | Should -Be 'GUASTO PREVISTO dallo SMART; SSD: usura 1%'
            Format-HDSmart | Should -Be ''
        }
    }

    It 'BitLocker (per il foglio principale), BIOS, monitor (con la marca) e stampanti vere' {
        $script:Det.BitLocker | Should -Be 'C: attiva'
        $script:Det.Dettagli.Bios | Should -Be '1.15.0 (12/03/2022)'
        $script:Det.Dettagli.Monitor | Should -Be 'DELL P2419H (S/N CFV9N93, 2020); Samsung S24F350 (2018)'
        $script:Det.Dettagli.Stampanti | Should -Be 'HP LaserJet M404 [192.168.1.50] (predefinita)'
    }

    It 'Office, attivazione di Windows, antivirus, accesso remoto' {
        $script:Det.Dettagli.Office | Should -Be 'Microsoft 365 Apps for business (16.0.17126)'
        $script:Det.Dettagli.AttivazioneWindows | Should -Be 'Attivato (OEM, chiave nel BIOS)'
        $script:Det.Dettagli.Antivirus | Should -Be 'Windows Defender (attivo, aggiornato)'
        $script:Det.Dettagli.AccessoRemoto | Should -Be 'TeamViewer'
    }

    It 'elenco dei programmi senza componenti di sistema e librerie, con data italiana' {
        $nomi = @($script:Det.Software | ForEach-Object { $_.Nome })
        $nomi | Should -Be @('Adobe Acrobat Reader', 'Microsoft 365 Apps for business - it-it', 'Mozilla Firefox (x64 it)', 'TeamViewer')
        ($script:Det.Software | Where-Object { $_.Nome -like 'Mozilla*' }).Installato | Should -Be '05/10/2024'
        $script:Det.Dettagli.Programmi | Should -Be '4'
    }

    It 'casi particolari: schermo del portatile, KMS, antivirus spento, Desktop remoto attivo' {
        InModuleScope HelpDeskToolkit {
            Format-HDMonitor @([pscustomobject]@{ Produttore = 'LGD'; Nome = ''; Seriale = '0'; Anno = 2021 }) | Should -Be 'Schermo integrato'
            Get-HDAttivazioneWindows @([pscustomobject]@{ Description = 'Windows(R) Operating System, VOLUME_KMSCLIENT channel'; LicenseStatus = 1 }) | Should -Be 'Attivato (Volume KMS)'
            Get-HDAttivazioneWindows @([pscustomobject]@{ Description = 'Windows(R) Operating System, RETAIL channel'; LicenseStatus = 0 }) | Should -Be 'Non attivato (Retail)'
            Get-HDAntivirusInstallati @([pscustomobject]@{ displayName = 'Windows Defender'; productState = 393472 }) | Should -Be 'Windows Defender (non attivo, aggiornato)'
            Get-HDAccessoRemoto -Programmi @([pscustomobject]@{ DisplayName = 'AnyDesk' }, [pscustomobject]@{ DisplayName = 'Supremo' }) -DesktopRemoto 0 | Should -Be 'AnyDesk; Supremo; Desktop remoto di Windows attivo'
            Format-HDBitLocker -Shell @([pscustomobject]@{ Unita = 'C:'; Valore = 2 }) | Should -Be 'C: non attiva'
        }
    }

    It 'su un PC che non ha dato nessun dato in più non si blocca' {
        $vuoto = InModuleScope HelpDeskToolkit { ConvertTo-HDDettagliPC -Grezzo ([pscustomobject]@{ NomePC = 'X' }) }
        $vuoto.Software | Should -BeNullOrEmpty
        $vuoto.Dettagli.Programmi | Should -Be ''
        $vuoto.Dettagli.Windows11 | Should -BeLike 'Da verificare*'
    }

    It 'il record del censimento porta BitLocker, Dettagli e Software' {
        $r = InModuleScope HelpDeskToolkit -Parameters @{ G = (New-GrezzoCompleto) } { param($G) ConvertTo-HDRecordCensimento -Grezzo $G }
        $r.BitLocker | Should -Be 'C: attiva'
        $r.Dettagli.TPM | Should -Be '2.0'
        @($r.Software).Count | Should -Be 4
    }
}

Describe 'Foglio Software: una riga per programma' {
    It 'riassume i PC: tutti, tutti tranne, elenco; e le versioni' {
        InModuleScope HelpDeskToolkit {
            $righe = foreach ($pc in 'PC-1', 'PC-2', 'PC-3', 'PC-4', 'PC-5') {
                $programmi = @([pscustomobject]@{ Nome = 'Google Chrome'; Versione = '131.0'; Editore = 'Google LLC' })
                if ($pc -ne 'PC-4') { $programmi += [pscustomobject]@{ Nome = 'Adobe Acrobat Reader'; Versione = '24.003'; Editore = 'Adobe' } }
                if ($pc -eq 'PC-2') { $programmi[0].Versione = '130.0'; $programmi += [pscustomobject]@{ Nome = 'AutoCAD 2024'; Versione = '24.3'; Editore = 'Autodesk' } }
                [pscustomobject]@{ Numero = 1; Valori = @{ Hostname = $pc; Programmi = (ConvertTo-HDTestoProgrammi $programmi) } }
            }
            $foglio = @(ConvertTo-HDRigheSoftware $righe)
            $foglio.Count | Should -Be 3
            ($foglio | Where-Object { $_[0] -eq 'Google Chrome' }) | Should -Be @('Google Chrome', 'Google LLC', '131.0 (4 PC); 130.0: PC-2', '5', 'Tutti')
            ($foglio | Where-Object { $_[0] -eq 'Adobe Acrobat Reader' }) | Should -Be @('Adobe Acrobat Reader', 'Adobe', '24.003', '4', 'Tutti tranne PC-4')
            ($foglio | Where-Object { $_[0] -eq 'AutoCAD 2024' }) | Should -Be @('AutoCAD 2024', 'Autodesk', '24.3', '1', 'PC-2')
            $foglio[0][0] | Should -Be 'Adobe Acrobat Reader'
        }
    }

    It 'pochi PC: elenco dei nomi; un elenco troppo lungo per una cella si accorcia' {
        InModuleScope HelpDeskToolkit {
            Format-HDElencoPC -Con @('PC-2', 'PC-1') -Tutti @('PC-1', 'PC-2') | Should -Be 'PC-1, PC-2'
            Format-HDElencoPC -Con @('PC-1') -Tutti @('PC-1', 'PC-2', 'PC-3', 'PC-4') | Should -Be 'PC-1'
            $tanti = @(1..2000 | ForEach-Object { [pscustomobject]@{ Nome = "Programma con un nome lungo numero $_"; Versione = '1.0.0.1234'; Editore = 'Editore con un nome lungo' } })
            $testo = ConvertTo-HDTestoProgrammi $tanti
            $testo.Length | Should -BeLessOrEqual $script:HDMassimoCella
            @(ConvertFrom-HDTestoProgrammi $testo)[0].Nome | Should -Be 'Programma con un nome lungo numero 1'
        }
    }
}

Describe 'Fogli Dettagli e Software nell''Excel' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:CartellaCensimentoTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        $script:File = (New-Censimento -Cliente 'Dettagli' -Confirm:$false).FullName
    }

    It 'un Excel nuovo ha i fogli Dettagli, Software e DatiSoftware (nascosto) e la colonna BitLocker' {
        @(Get-Fogli $script:File | ForEach-Object { $_.Nome }) | Should -Be @('Censimento', 'Dettagli', 'Software', 'DatiSoftware')
    }

    It 'censisce due PC: Dettagli una riga per PC, Software una riga per programma; poi aggiorna senza doppioni' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        $e = Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null
        $e.Scritto | Should -BeTrue
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario {
            New-GrezzoCompleto -Nome 'PC-RECEPTION' -Serie 'SER-002' -Mac '00:11:22:33:44:02' -Programmi @(
                [pscustomobject]@{ DisplayName = 'Mozilla Firefox (x64 it)'; DisplayVersion = '130.0'; Publisher = 'Mozilla' }
                [pscustomobject]@{ DisplayName = '7-Zip 24.08 (x64)'; DisplayVersion = '24.08'; Publisher = 'Igor Pavlov' }
            )
        }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        $fogli = Get-Fogli $script:File
        $censimento = @(Get-RigheFoglio $fogli 'Censimento')
        $censimento.Count | Should -Be 2
        $censimento[0][20] | Should -Be 'C: attiva'
        $dettagli = @(Get-RigheFoglio $fogli 'Dettagli')
        $dettagli.Count | Should -Be 2
        $dettagli[0][4] | Should -Be '2.0'
        $software = @(Get-RigheFoglio $fogli 'Software')
        @($software | ForEach-Object { $_[1] }) | Should -Be @('7-Zip 24.08 (x64)', 'Adobe Acrobat Reader', 'Microsoft 365 Apps for business - it-it', 'Mozilla Firefox (x64 it)', 'TeamViewer')
        ($software | Where-Object { $_[1] -like 'Mozilla*' })[3] | Should -Be '131.0: PC-AMMINISTRAZIONE; 130.0: PC-RECEPTION'
        ($software | Where-Object { $_[1] -like 'Mozilla*' })[4] | Should -Be '2'
        ($software | Where-Object { $_[1] -like 'Mozilla*' })[5] | Should -Be 'PC-AMMINISTRAZIONE, PC-RECEPTION'
        ($software | Where-Object { $_[1] -like 'Adobe*' })[5] | Should -Be 'PC-AMMINISTRAZIONE'
        @(Get-RigheFoglio $fogli 'DatiSoftware').Count | Should -Be 2

        # di nuovo il primo PC, senza Adobe: sparisce dal riepilogo, nessuna riga doppia
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario {
            New-GrezzoCompleto -Programmi @([pscustomobject]@{ DisplayName = 'Mozilla Firefox (x64 it)'; DisplayVersion = '132.0'; Publisher = 'Mozilla' })
        }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        $fogli = Get-Fogli $script:File
        $software = @(Get-RigheFoglio $fogli 'Software')
        @($software | ForEach-Object { $_[1] }) | Should -Be @('7-Zip 24.08 (x64)', 'Mozilla Firefox (x64 it)')
        @(Get-RigheFoglio $fogli 'DatiSoftware').Count | Should -Be 2
        @(Get-RigheFoglio $fogli 'Dettagli').Count | Should -Be 2
    }

    It 'a un Excel di una versione precedente aggiunge la colonna BitLocker in fondo e i fogli mancanti' {
        $vecchio = Join-Path $script:CartellaCensimentoTest 'Censimento_Vecchio.xlsx'
        InModuleScope HelpDeskToolkit -Parameters @{ P = $vecchio } { param($P) New-HDXlsx -Percorso $P -NomeFoglio 'Censimento' -Colonne @($script:HDColonneCensimento | Where-Object { $_.Chiave -ne 'BitLocker' }) }
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto -Nome 'PC-VECCHIO' -Serie 'SER-OLD' }
        $e = Add-PCCensimento -Censimento $vecchio -NonChiedere -Confirm:$false -WarningVariable avvisi 6>$null
        $e.Scritto | Should -BeTrue
        @($avvisi) | Should -BeNullOrEmpty
        $fogli = Get-Fogli $vecchio
        @($fogli | ForEach-Object { $_.Nome }) | Should -Be @('Censimento', 'Dettagli', 'Software', 'DatiSoftware')
        $principale = $fogli | Where-Object Nome -eq 'Censimento'
        $principale.Righe[1][20] | Should -Be 'Locazione'
        $principale.Righe[1][21] | Should -Be 'BitLocker'
        $principale.Righe[2][21] | Should -Be 'C: attiva'
        ($fogli | Where-Object Nome -eq 'Dettagli').Righe[1][3] | Should -Be 'Pronto per Windows 11'
        @(Get-RigheFoglio $fogli 'Dettagli')[0][1] | Should -Be 'PC-VECCHIO'
        @(Get-RigheFoglio $fogli 'Software').Count | Should -Be 4
    }

    It 'un foglio Dettagli di prima riceve le colonne nuove (es. SMART) in fondo' {
        $prima = Join-Path $script:CartellaCensimentoTest 'Censimento_Prima.xlsx'
        InModuleScope HelpDeskToolkit -Parameters @{ P = $prima } {
            param($P)
            New-HDXlsx -Percorso $P -NomeFoglio 'Censimento' -Colonne $script:HDColonneCensimento -AltriFogli @(@{ Nome = 'Dettagli'; Colonne = @($script:HDColonneDettagli | Where-Object { $_.Chiave -ne 'Smart' }) })
        }
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        [void](Add-PCCensimento -Censimento $prima -NonChiedere -Confirm:$false 6>$null)
        $d = (Get-Fogli $prima | Where-Object Nome -eq 'Dettagli').Righe
        $ultima = InModuleScope HelpDeskToolkit { $script:HDColonneDettagli.Count }
        $d[1][$ultima] | Should -Be 'SMART dei dischi'
        $d[2][$ultima] | Should -BeLike 'Samsung SSD 980*'
    }

    It 'senza diritti di amministratore lo SMART dice come leggerlo, ma non cancella quello già letto' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { $g = New-GrezzoCompleto; $g.Smart = $null; $g }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        @(Get-RigheFoglio (Get-Fogli $script:File) 'Dettagli')[0][8] | Should -BeLike 'Avvia il toolkit come amministratore*'
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { $g = New-GrezzoCompleto; $g.Smart = $null; $g }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        @(Get-RigheFoglio (Get-Fogli $script:File) 'Dettagli')[0][8] | Should -BeLike 'Samsung SSD 980*'
    }

    It 'non tocca un foglio "Software" che contiene altro' {
        $altro = Join-Path $script:CartellaCensimentoTest 'Censimento_Altro.xlsx'
        InModuleScope HelpDeskToolkit -Parameters @{ P = $altro } {
            param($P)
            New-HDXlsx -Percorso $P -NomeFoglio 'Censimento' -Colonne $script:HDColonneCensimento -AltriFogli @(@{ Nome = 'Software'; Colonne = @([pscustomobject]@{ Titolo = 'Licenze acquistate'; Larghezza = 20 }) })
        }
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        $e = Add-PCCensimento -Censimento $altro -NonChiedere -Confirm:$false -WarningVariable avvisi -WarningAction SilentlyContinue 6>$null
        $e.Scritto | Should -BeTrue
        ($avvisi -join ' ') | Should -BeLike "*foglio 'Software'*"
        $fogli = Get-Fogli $altro
        ($fogli | Where-Object Nome -eq 'Software').Righe.Count | Should -Be 1
        @(Get-RigheFoglio $fogli 'Dettagli').Count | Should -Be 1
    }

    It 'la finestra può passare i dati già letti (-Record) senza rileggere il PC' {
        $r = InModuleScope HelpDeskToolkit -Parameters @{ G = (New-GrezzoCompleto -Nome 'PC-FINESTRA') } { param($G) ConvertTo-HDRecordCensimento -Grezzo $G }
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { throw 'non doveva rileggere' }
        $e = Add-PCCensimento -Censimento $script:File -Record $r -NonChiedere -Confirm:$false 6>$null
        $e.Scritto | Should -BeTrue
        Should -Invoke -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario -Times 0
        @(Get-RigheFoglio (Get-Fogli $script:File) 'Dettagli')[0][1] | Should -Be 'PC-FINESTRA'
    }
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    . (Join-Path $PSScriptRoot 'DatiDiProva.ps1')

    # PC censito come lo vede il riepilogo: valori del foglio principale e del foglio Dettagli
    function New-PcRiepilogo {
        param([string]$Nome, [hashtable]$Valori = @{}, [hashtable]$Dettagli = @{}, [string]$Utilizzatore = '')
        $v = @{ Hostname = $Nome; Utilizzatore = $Utilizzatore }
        foreach ($k in $Valori.Keys) { $v[$k] = $Valori[$k] }
        [pscustomobject]@{ Nome = $Nome; Utilizzatore = $Utilizzatore; Valori = $v; Dettagli = $Dettagli }
    }
    # "PC: frase" per ogni PC di un controllo
    function Get-Fatti($Punti, [string]$Punto) {
        $p = @($Punti | Where-Object { $_.Punto -eq $Punto }) | Select-Object -First 1
        if (-not $p) { return @() }
        @($p.PC | ForEach-Object { '{0}: {1}' -f $_.Nome, $_.Fatto })
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Riepilogo: problemi del parco PC' {
    It 'Windows senza aggiornamenti: <Atteso>' -TestCases @(
        @{ OS = 'Microsoft Windows 10 Pro'; Versione = '22H2 (19045.5011)'; Atteso = 'Windows 10' }
        @{ OS = 'Microsoft Windows 10 Enterprise LTSC'; Versione = '21H2'; Atteso = '' }
        @{ OS = 'Microsoft Windows 11 Pro'; Versione = '23H2 (22631.4317)'; Atteso = 'Windows 11 23H2' }
        @{ OS = 'Microsoft Windows 11 Pro'; Versione = '24H2 (26100.2033)'; Atteso = 'Windows 11 24H2' }
        @{ OS = 'Microsoft Windows 11 Enterprise'; Versione = '23H2 (22631.4317)'; Atteso = '' }
        @{ OS = 'Microsoft Windows 11 Pro'; Versione = '25H2 (26200.1)'; Atteso = '' }
        @{ OS = 'Microsoft Windows 11 Pro'; Versione = '26H2 (28000.1)'; Atteso = '' }
        @{ OS = 'Microsoft Windows 7 Professional'; Versione = ''; Atteso = 'Windows 7' }
        @{ OS = 'Microsoft Windows Server 2019 Standard'; Versione = '1809'; Atteso = '' }
    ) {
        param($OS, $Versione, $Atteso)
        InModuleScope HelpDeskToolkit -Parameters @{ O = $OS; V = $Versione } {
            param($O, $V)
            Get-HDWindowsSenzaAggiornamenti -SistemaOperativo $O -Versione $V -Adesso ([datetime]'2026-11-07')
        } | Should -Be $Atteso
    }

    It 'dice per ogni PC il problema preciso, in una frase breve' {
        $vecchio = New-PcRiepilogo 'PC-VECCHIO' -Utilizzatore 'Mario Rossi' -Valori @{
            OS = 'Microsoft Windows 10 Pro'; OSVersion = '22H2 (19045.5011)'; Cynet = 'No'; BitLocker = 'C: non attiva'
            Memory = '4 GB DDR3 SODIMM'; FreeDiskPct = '(C:) 8 GB 3%'
        } -Dettagli @{
            Windows11 = 'Sì (prima attiva Secure Boot nel BIOS)'; Smart = 'GUASTO PREVISTO dallo SMART; ST500DM002: 41 °C, ATTENZIONE 3 errori non corretti'
            DischiFisici = 'ST500DM002 (HDD SATA, 466 GB, ATTENZIONE)'
            Batteria = '44% (22 di 50 Wh) - da sostituire'; Amministratori = 'Administrator (disattivato), mario (utente del PC)'
            Aggiornamento = '01/06/2026 (KB5060842)'; Rilevato = '02/11/2026 10:00'; Office = 'Office Professional Plus 2016'
            AccessoRemoto = 'TeamViewer; Desktop remoto di Windows attivo'; Bios = 'A12 (01/03/2017)'; AttivazioneWindows = 'Non attivato'
            Antivirus = 'Windows Defender (non attivo, firme non aggiornate)'
        }
        $nuovo = New-PcRiepilogo 'PC-NUOVO' -Valori @{
            OS = 'Microsoft Windows 11 Pro'; OSVersion = '25H2 (26200.7019)'; Cynet = 'Sì'; BitLocker = 'C: attiva; D: non attiva'
            Memory = '16 GB DDR5'; FreeDiskPct = '(C:) 200 GB 45%'
        } -Dettagli @{
            Windows11 = 'Già Windows 11'; Smart = 'Samsung SSD 990: usura 1%, 35 °C, nessun errore'; Batteria = 'Nessuna'
            Amministratori = 'Administrator (disattivato), STUDIO\Domain Admins'; Aggiornamento = '14/10/2026 (KB5066835)'
            Rilevato = '05/11/2026 09:00'; Office = 'Microsoft 365 Apps for business (16.0.19127)'; AccessoRemoto = 'Nessuno'
            Bios = '1.20.0 (15/06/2025)'; AttivazioneWindows = 'Attivato (OEM, chiave nel BIOS)'; Antivirus = 'Windows Defender (attivo, aggiornato)'
        }
        $dimenticato = New-PcRiepilogo 'PC-MAGAZZINO' -Valori @{
            OS = 'Microsoft Windows 11 Pro'; OSVersion = '23H2 (22631.4317)'; Cynet = 'Sì'; FreeDiskPct = '(C:) 15 GB 12%'
        } -Dettagli @{ Rilevato = '10/01/2026 15:00'; Batteria = 'Nessuna' }

        $punti = @(InModuleScope HelpDeskToolkit -Parameters @{ P = @($vecchio, $nuovo, $dimenticato) } {
                param($P)
                Get-HDPuntiRiepilogo -PC $P -Soglie (Get-HDConfig).Soglie -Adesso ([datetime]'2026-11-07')
            })
        $punti[0].Priorita | Should -Be 'Urgente'
        Get-Fatti $punti 'Disco a rischio di guasto' | Should -Be @('PC-VECCHIO: Disco a rischio di guasto: ST500DM002')
        Get-Fatti $punti 'Windows senza aggiornamenti di sicurezza' | Should -Be @(
            'PC-VECCHIO: Windows 10 senza aggiornamenti di sicurezza (compatibile con Windows 11)'
            'PC-MAGAZZINO: Windows 11 23H2 senza aggiornamenti di sicurezza'
        )
        Get-Fatti $punti 'Antivirus spento' | Should -Be @('PC-VECCHIO: Antivirus spento: Windows Defender')
        Get-Fatti $punti 'Cynet non installato' | Should -Be @('PC-VECCHIO: Cynet non installato')
        Get-Fatti $punti 'Disco pieno' | Should -Be @('PC-VECCHIO: Disco C: pieno (8 GB liberi)')
        Get-Fatti $punti 'Windows non attivato' | Should -Be @('PC-VECCHIO: Windows non attivato')
        Get-Fatti $punti 'Aggiornamenti di Windows fermi' | Should -Be @('PC-VECCHIO: Aggiornamenti di Windows fermi dal 01/06/2026')
        Get-Fatti $punti 'Disco quasi pieno' | Should -Be @('PC-MAGAZZINO: Disco C: quasi pieno (15 GB liberi)')
        Get-Fatti $punti 'BitLocker non attivo' | Should -Be @('PC-VECCHIO: BitLocker non attivo (portatile)')
        Get-Fatti $punti 'Utente amministratore del PC' | Should -Be @("PC-VECCHIO: L'utente mario è amministratore")
        Get-Fatti $punti 'Desktop remoto (RDP) attivo' | Should -Be @('PC-VECCHIO: Desktop remoto (RDP) attivo')
        Get-Fatti $punti 'Programmi di accesso remoto' | Should -Be @('PC-VECCHIO: TeamViewer installato')
        Get-Fatti $punti 'Office non più supportato' | Should -Be @('PC-VECCHIO: Office 2016 non più supportato')
        Get-Fatti $punti 'Batteria usurata' | Should -Be @('PC-VECCHIO: Batteria al 44%, da sostituire')
        Get-Fatti $punti 'Poca RAM' | Should -Be @('PC-VECCHIO: Solo 4 GB di RAM')
        Get-Fatti $punti 'BIOS vecchio' | Should -Be @('PC-VECCHIO: BIOS del 2017')
        Get-Fatti $punti 'Non censito da tempo' | Should -Be @('PC-MAGAZZINO: Non censito dal 10/01/2026')
        # il PC nuovo e in ordine non compare in nessun controllo
        @($punti | Where-Object { @($_.PC | Where-Object { $_.Nome -eq 'PC-NUOVO' }).Count -gt 0 }) | Should -BeNullOrEmpty

        # foglio Riepilogo: una riga per PC, prima chi sta peggio, urgente e da controllare separati
        $righe = @(InModuleScope HelpDeskToolkit -Parameters @{ P = @($vecchio, $nuovo, $dimenticato); Q = $punti } {
                param($P, $Q)
                ConvertTo-HDRigheRiepilogo -PerPc @(Get-HDRiepilogoPerPc -PC $P -Punti $Q) -Adesso ([datetime]'2026-11-07 14:30')
            })
        $righe.Count | Should -Be 4
        @($righe | Select-Object -First 3 | ForEach-Object { $_[0] }) | Should -Be @('PC-VECCHIO', 'PC-MAGAZZINO', 'PC-NUOVO')
        $righe[0][1] | Should -Be 'Mario Rossi'
        $righe[0][2] | Should -Be 'Disco a rischio di guasto: ST500DM002; Windows 10 senza aggiornamenti di sicurezza (compatibile con Windows 11); Antivirus spento: Windows Defender; Cynet non installato; Disco C: pieno (8 GB liberi); Windows non attivato'
        $righe[0][3] | Should -Be "Aggiornamenti di Windows fermi dal 01/06/2026; BitLocker non attivo (portatile); L'utente mario è amministratore; Desktop remoto (RDP) attivo; TeamViewer installato; Office 2016 non più supportato; Batteria al 44%, da sostituire; Solo 4 GB di RAM; BIOS del 2017"
        $righe[1][2] | Should -Be 'Windows 11 23H2 senza aggiornamenti di sicurezza'
        $righe[1][3] | Should -Be 'Disco C: quasi pieno (15 GB liberi); Non censito dal 10/01/2026'
        $righe[2][2] | Should -Be ''
        $righe[3][0] | Should -Be 'Calcolato il 07/11/2026 14:30 su 3 PC censiti: si aggiorna da solo a ogni censimento, non scriverci dentro.'
    }

    It 'Cynet si segnala solo dove è in uso; i dispositivi visti solo dalla scansione non contano' {
        $righe = [System.Collections.Generic.List[object]]::new()
        $righe.Add([pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; OS = 'Microsoft Windows 11 Pro'; OSVersion = '25H2'; Cynet = 'No' } })
        $righe.Add([pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'STAMPANTE'; IndirizzoIP = '192.168.1.50' } })
        $righe.Add([pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'PC-02'; SerialNumber = 'S2' } })
        $dettagli = @([pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-02'; SerialNumber = 'S2'; Rilevato = '01/11/2026 10:00'; Batteria = 'Nessuna' } })
        InModuleScope HelpDeskToolkit -Parameters @{ R = $righe; D = $dettagli } {
            param($R, $D)
            $pc = @(Get-HDPcCensiti -Righe @($R | Where-Object { $_ }) -Dettagli $D)
            @($pc | ForEach-Object { $_.Nome }) | Should -Be @('PC-01', 'PC-02')
            $pc[1].Dettagli['Batteria'] | Should -Be 'Nessuna'
            @(Get-HDPuntiRiepilogo -PC $pc -Adesso ([datetime]'2026-11-07') | Where-Object { $_.Punto -eq 'Cynet non installato' }) | Should -BeNullOrEmpty
        }
    }

    It 'nome del cliente dal file; dischi a rischio anche senza nome' {
        InModuleScope HelpDeskToolkit {
            Get-HDNomeCliente 'C:\Chiavetta\Censimento\Censimento_Studio_Rossi.xlsx' | Should -Be 'Studio Rossi'
            Get-HDNomeCliente 'Censimento_Studio-Rossi (2).xlsx' | Should -Be 'Studio Rossi'
            Get-HDNomeCliente 'Censimento.xlsx' | Should -Be ''
            Get-HDDischiARischio -Smart 'GUASTO PREVISTO dallo SMART; SSD: usura 2%, nessun errore' -DischiFisici 'SSD (SSD NVMe, 500 GB, OK)' | Should -Be ''
            Get-HDDischiARischio -Smart 'SSD: usura 2%, nessun errore' -DischiFisici 'SSD (SSD NVMe, 500 GB, OK)' | Should -BeNullOrEmpty
            Join-HDElenco @('TeamViewer', 'AnyDesk', 'RustDesk') | Should -Be 'TeamViewer, AnyDesk e RustDesk'
        }
    }
}

Describe 'Storico dei cambiamenti' {
    It 'una frase per cambiamento; niente primo censimento, build, celle vuote' {
        InModuleScope HelpDeskToolkit {
            $nuovaTabella = {
                param($Righe)
                $l = [System.Collections.Generic.List[object]]::new()
                foreach ($r in $Righe) { $l.Add($r) }
                [pscustomobject]@{ Righe = $l }
            }
            $principale = & $nuovaTabella @(
                [pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; OS = 'Microsoft Windows 11 Pro'; OSVersion = '24H2 (26100.2033)'; Memory = '8 GB DDR4'; NomeUtente = 'mario'; BitLocker = 'C: non attiva' } }
                [pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'PC-02'; IndirizzoIP = '192.168.1.22'; Manufacturer = 'ASUSTek COMPUTER INC.' } }
            )
            $tabella = [pscustomobject]@{
                Righe        = $principale.Righe
                Dettagli     = (& $nuovaTabella @([pscustomobject]@{ Numero = 2; Valori = @{
                                Hostname = 'PC-01'; SerialNumber = 'S1'; Rilevato = '01/10/2026 09:00'; Office = 'Microsoft 365 Apps for business (16.0.18000)'
                                Monitor = 'DELL P2419H (S/N A1, 2020)'; Amministratori = 'Administrator (disattivato), mario (utente del PC)'
                                DischiFisici = 'WDC WD10EZEX (HDD SATA, 1000 GB, OK)'
                            }
                        }))
                DatiSoftware = (& $nuovaTabella @([pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; Programmi = "7-Zip 24.08 (x64)`t24.08`tIgor Pavlov`nAdobe Acrobat Reader`t24.1`tAdobe`nTeamViewer`t15.50`tTeamViewer" } }))
            }
            $foto = Get-HDFotoCensimento $tabella

            # nuovo censimento: più RAM, BitLocker attivato, build nuova (non conta), stesso Office a parte il numero
            $pc1 = $tabella.Righe[0].Valori
            $pc1.Memory = '16 GB DDR4 (2 x 8 GB)'
            $pc1.OSVersion = '24H2 (26100.4061)'
            $pc1.BitLocker = 'C: attiva'
            $pc1.Dominio = 'studio.local'
            $d1 = $tabella.Dettagli.Righe[0].Valori
            $d1.Rilevato = '05/11/2026 10:00'
            $d1.Office = 'Microsoft 365 Apps for business (16.0.19127)'
            $d1.Monitor = 'DELL P2419H (S/N A1, 2020); LG 27UL500 (2021)'
            $d1.Amministratori = 'Administrator (disattivato), mario (utente del PC), luca'
            $d1.DischiFisici = 'WDC WD10EZEX (HDD SATA, 1000 GB, ATTENZIONE)'
            $tabella.DatiSoftware.Righe[0].Valori.Programmi = "7-Zip 24.09 (x64)`t24.09`tIgor Pavlov`nTeamViewer`t15.60`tTeamViewer`nAnyDesk`t9.0`tAnyDesk"
            # la riga arrivata dalla scansione ora è censita (e il produttore cambia nome), e c'è un PC nuovo
            $tabella.Righe[1].Valori.OS = 'Microsoft Windows 10 Pro'
            $tabella.Righe[1].Valori.Manufacturer = 'ASUS'
            $tabella.Righe.Add([pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'PC-03'; OS = 'Microsoft Windows 11 Pro' } })

            $voci = @(Get-HDCambiamenti -Foto $foto -Tabella $tabella -Adesso ([datetime]'2026-11-07 12:00'))
            @($voci | ForEach-Object { '{0}|{1}|{2}' -f $_.Data, $_.Hostname, $_.Cambiamento }) | Should -Be @(
                '05/11/2026 10:00|PC-01|RAM: 8 GB DDR4 → 16 GB DDR4 (2 x 8 GB)'
                '05/11/2026 10:00|PC-01|BitLocker: C: non attiva → C: attiva'
                '05/11/2026 10:00|PC-01|Dischi fisici: WDC WD10EZEX (HDD SATA, 1000 GB, OK) → WDC WD10EZEX (HDD SATA, 1000 GB, ATTENZIONE)'
                '05/11/2026 10:00|PC-01|Monitor: aggiunto LG 27UL500 (2021)'
                '05/11/2026 10:00|PC-01|Amministratori: aggiunto luca'
                '05/11/2026 10:00|PC-01|Programmi installati: AnyDesk'
                '05/11/2026 10:00|PC-01|Programmi disinstallati: Adobe Acrobat Reader'
            )

            # in fondo al foglio Storico, una riga per voce
            $storico = ConvertTo-HDTabellaFoglio -Fogli @() -Nome 'Storico' -Colonne $script:HDColonneStorico
            $modifiche = @{}
            Add-HDStoricoATabella -Tabella $storico -Voci $voci -Modifiche $modifiche
            @($modifiche.Keys | Sort-Object) | Should -Be @(2, 3, 4, 5, 6, 7, 8)
            $modifiche[2].Cambiamento | Should -BeLike 'RAM:*'
            $modifiche[2].Keys | Sort-Object | Should -Be @('Cambiamento', 'Data', 'Hostname', 'SerialNumber')
        }
    }
}

Describe 'Pagina del riepilogo per il cliente' {
    It 'sintesi per problema, PC per PC con i fatti, numeri, cambiamenti; niente consigli generici, testo sicuro' {
        $html = InModuleScope HelpDeskToolkit {
            $pc = @(
                [pscustomobject]@{ Nome = 'PC-01'; Utilizzatore = 'Anna <b>Bianchi</b>'; Valori = @{ OS = 'Microsoft Windows 10 Pro'; OSVersion = '22H2 (19045.1)'; Memory = '8 GB DDR4'; Manufacturer = 'Dell Inc.'; Model = 'OptiPlex 7090'; BitLocker = 'C: non attiva' }; Dettagli = @{ Windows11 = 'Sì'; Batteria = 'Nessuna'; Rilevato = '02/11/2026 10:00'; AccessoRemoto = 'Desktop remoto di Windows attivo' } }
                [pscustomobject]@{ Nome = 'PC-02'; Utilizzatore = ''; Valori = @{ OS = 'Microsoft Windows 11 Pro'; OSVersion = '25H2 (26200.1)'; Memory = '16 GB' }; Dettagli = @{ Batteria = '90% (45 di 50 Wh)'; Rilevato = '05/11/2026 10:00'; Office = 'Microsoft 365 Apps for business' } }
            )
            $punti = @(Get-HDPuntiRiepilogo -PC $pc -Adesso ([datetime]'2026-11-07'))
            $cambi = @([pscustomobject]@{ Data = '05/11/2026 10:00'; Hostname = 'PC-02'; Cambiamento = 'RAM: 8 GB → 16 GB' }, [pscustomobject]@{ Data = '01/01/2025 10:00'; Hostname = 'PC-01'; Cambiamento = 'Vecchio cambiamento' })
            ConvertTo-HDHtmlRiepilogo -Cliente 'Studio Rossi' -PC $pc -Punti $punti -Cambiamenti $cambi -File 'Censimento_Studio_Rossi.xlsx' -Adesso ([datetime]'2026-11-07 15:00')
        }
        # le lettere accentate diventano entità (pi&#249;): si controlla il testo decodificato
        $testo = [System.Net.WebUtility]::HtmlDecode($html)
        $testo | Should -Match '<h1>Studio Rossi</h1>'
        $testo | Should -Match '>1 PC con problemi urgenti<'
        $testo | Should -Match '1 PC su 2 senza problemi'
        $testo | Should -Match '<h2>In sintesi</h2>'
        $testo | Should -Match 'Windows senza aggiornamenti di sicurezza</td><td class="num">1 PC'
        $testo | Should -Match '<h2>PC per PC</h2>'
        $testo | Should -Match 'Windows 10 senza aggiornamenti di sicurezza \(compatibile con Windows 11\)'
        $testo | Should -Match 'BitLocker non attivo'
        $testo | Should -Match 'Desktop remoto \(RDP\) attivo'
        $testo | Should -Match 'Senza problemi: PC-02'
        $testo | Should -Match 'Portatili / fissi</dt><dd>1 / 1'
        $testo | Should -Match 'RAM: 8 GB → 16 GB'
        $testo | Should -Not -Match 'Vecchio cambiamento'
        $testo | Should -Not -Match 'Cosa fare'
        $html | Should -Match 'Anna &lt;b&gt;Bianchi&lt;/b&gt;'
        $html | Should -Not -Match '<b>Bianchi'
        $testo | Should -Not -Match 'tua azienda'
    }
}

Describe 'Riepilogo e Storico nell''Excel' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:CartellaCensimentoTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        Get-ChildItem -LiteralPath $script:CartellaReportTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        $script:File = (New-Censimento -Cliente 'Studio Rossi' -Confirm:$false).FullName
    }

    It 'a ogni censimento il Riepilogo si ricalcola (una riga per PC) e lo Storico registra solo i cambiamenti' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        (Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null).Scritto | Should -BeTrue
        $fogli = Get-Fogli $script:File
        $riepilogo = @(Get-RigheFoglio $fogli 'Riepilogo')
        $riepilogo.Count | Should -Be 2
        $riepilogo[0][1] | Should -Be 'PC-AMMINISTRAZIONE'
        $riepilogo[0][3] | Should -Be 'Disco a rischio di guasto: WDC WD10EZEX; Windows 10 senza aggiornamenti di sicurezza (compatibile con Windows 11)'
        $riepilogo[0][4] | Should -Match "L'utente STUDIO\\giulia è amministratore"
        $riepilogo[0][4] | Should -Match 'TeamViewer installato'
        $riepilogo[1][1] | Should -BeLike 'Calcolato il * su 1 PC censito*'
        @(Get-RigheFoglio $fogli 'Storico') | Should -BeNullOrEmpty

        # di nuovo lo stesso PC con più RAM e un programma in più
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario {
            $g = New-GrezzoCompleto
            $g.Memoria = @([pscustomobject]@{ Capacity = 16GB }, [pscustomobject]@{ Capacity = 16GB })
            $g.Programmi = @($g.Programmi) + @([pscustomobject]@{ DisplayName = 'AnyDesk'; DisplayVersion = '9.0'; Publisher = 'AnyDesk Software GmbH' })
            $g
        }
        (Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null).Scritto | Should -BeTrue
        $fogli = Get-Fogli $script:File
        $storico = @(Get-RigheFoglio $fogli 'Storico')
        @($storico | ForEach-Object { $_[4] }) | Should -Be @('RAM: 16 GB (2 x 8 GB) → 32 GB (2 x 16 GB)', 'Accesso remoto: aggiunto AnyDesk', 'Programmi installati: AnyDesk')
        $storico[0][2] | Should -Be 'PC-AMMINISTRAZIONE'
        (@(Get-RigheFoglio $fogli 'Riepilogo')[0])[4] | Should -Match 'TeamViewer e AnyDesk installati'
    }

    It 'Export-RiepilogoCensimento crea la pagina nella cartella Report' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        $pagina = Export-RiepilogoCensimento -Censimento $script:File -Confirm:$false
        $pagina.DirectoryName | Should -Be $script:CartellaReportTest
        $pagina.Name | Should -BeLike 'Riepilogo_Studio-Rossi_*.html'
        $testo = [System.Net.WebUtility]::HtmlDecode((Get-Content -LiteralPath $pagina.FullName -Raw -Encoding UTF8))
        $testo | Should -Match '<h1>Studio Rossi</h1>'
        $testo | Should -Match 'PC-AMMINISTRAZIONE'
        $testo | Should -Match 'Disco a rischio di guasto: WDC WD10EZEX'
    }
}

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
    function Get-Punto($Punti, [string]$Punto) { @($Punti | Where-Object { $_.Punto -eq $Punto }) | Select-Object -First 1 }
    function Get-ElencoPunto($Punti, [string]$Punto) {
        $p = Get-Punto $Punti $Punto
        if (-not $p) { return @() }
        @($p.PC | ForEach-Object { if ($_.Dettaglio) { '{0} ({1})' -f $_.Nome, $_.Dettaglio } else { $_.Nome } })
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Riepilogo: punti da sistemare nel parco PC' {
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

    It 'trova i punti dai dati del foglio principale e di Dettagli, con il dettaglio di ogni PC' {
        $vecchio = New-PcRiepilogo 'PC-VECCHIO' -Utilizzatore 'Mario Rossi' -Valori @{
            OS = 'Microsoft Windows 10 Pro'; OSVersion = '22H2 (19045.5011)'; Cynet = 'No'; BitLocker = 'C: non attiva'
            Memory = '4 GB DDR3 SODIMM'; FreeDiskPct = '(C:) 8 GB 3%'
        } -Dettagli @{
            Windows11 = 'Sì (prima attiva Secure Boot nel BIOS)'; Smart = 'GUASTO PREVISTO dallo SMART; ST500: 41 °C'
            Batteria = '44% (22 di 50 Wh) - da sostituire'; Amministratori = 'Administrator (disattivato), mario (utente del PC)'
            Aggiornamento = '01/06/2026 (KB5060842)'; Rilevato = '02/11/2026 10:00'; Office = 'Office Professional Plus 2016'
            AccessoRemoto = 'TeamViewer'; Bios = 'A12 (01/03/2017)'; AttivazioneWindows = 'Non attivato'
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
        Get-ElencoPunto $punti 'Dischi a rischio di guasto' | Should -Be @('PC-VECCHIO')
        Get-ElencoPunto $punti 'Windows non riceve più aggiornamenti di sicurezza' | Should -Be @('PC-VECCHIO (Windows 10, compatibile con Windows 11)')
        Get-ElencoPunto $punti 'Versione di Windows 11 non più aggiornata' | Should -Be @('PC-MAGAZZINO (Windows 11 23H2)')
        Get-ElencoPunto $punti 'Antivirus assente o spento' | Should -Be @('PC-VECCHIO (Windows Defender (non attivo, firme non aggiornate))')
        Get-ElencoPunto $punti 'Cynet non installato' | Should -Be @('PC-VECCHIO')
        Get-ElencoPunto $punti 'Disco pieno' | Should -Be @('PC-VECCHIO (C: 8 GB liberi)')
        Get-ElencoPunto $punti 'Windows non attivato' | Should -Be @('PC-VECCHIO')
        Get-ElencoPunto $punti 'Aggiornamenti di Windows in ritardo' | Should -Be @('PC-VECCHIO (ultimo il 01/06/2026)')
        Get-ElencoPunto $punti 'Disco con poco spazio libero' | Should -Be @('PC-MAGAZZINO (C: 15 GB liberi)')
        Get-ElencoPunto $punti 'BitLocker non attivo' | Should -Be @('PC-VECCHIO (portatile)')
        Get-ElencoPunto $punti 'L''utente lavora come amministratore del PC' | Should -Be @('PC-VECCHIO')
        Get-ElencoPunto $punti 'Office non più supportato' | Should -Be @('PC-VECCHIO (Office 2016)')
        Get-ElencoPunto $punti 'Batteria del portatile usurata' | Should -Be @('PC-VECCHIO (44%)')
        Get-ElencoPunto $punti 'Poca memoria RAM' | Should -Be @('PC-VECCHIO (4 GB)')
        Get-ElencoPunto $punti 'Programmi di accesso remoto' | Should -Be @('PC-VECCHIO (TeamViewer)')
        Get-ElencoPunto $punti 'BIOS fermo da anni' | Should -Be @('PC-VECCHIO (BIOS del 2017)')
        Get-ElencoPunto $punti 'PC non censiti da tempo' | Should -Be @('PC-MAGAZZINO (ultimo censimento il 10/01/2026)')
        # il PC nuovo e in ordine non compare in nessun punto
        @($punti | Where-Object { @($_.PC | Where-Object { $_.Nome -eq 'PC-NUOVO' }).Count -gt 0 }) | Should -BeNullOrEmpty
        # l'utilizzatore resta per il report
        (Get-Punto $punti 'Disco pieno').PC[0].Utilizzatore | Should -Be 'Mario Rossi'
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

    It 'righe del foglio Riepilogo, nome del cliente dal file' {
        InModuleScope HelpDeskToolkit {
            $punto = [pscustomobject]@{ Priorita = 'Urgente'; Punto = 'Disco pieno'; CosaFare = 'Liberare spazio.'; PC = @([pscustomobject]@{ Nome = 'PC-01'; Utilizzatore = ''; Dettaglio = 'C: 3 GB liberi' }, [pscustomobject]@{ Nome = 'PC-02'; Utilizzatore = ''; Dettaglio = '' }) }
            $righe = @(ConvertTo-HDRigheRiepilogo -Punti @($punto) -TotalePC 3 -Adesso ([datetime]'2026-11-07 14:30'))
            $righe.Count | Should -Be 2
            $righe[0] | Should -Be @('Urgente', 'Disco pieno', '2', 'PC-01 (C: 3 GB liberi), PC-02', 'Liberare spazio.')
            $righe[1][1] | Should -Be 'Calcolato il 07/11/2026 14:30 su 3 PC censiti: si aggiorna da solo a ogni censimento.'
            @(ConvertTo-HDRigheRiepilogo -Punti @() -TotalePC 1)[0][1] | Should -Be 'Nessun punto da sistemare'
            Get-HDNomeCliente 'C:\Chiavetta\Censimento\Censimento_Studio_Rossi.xlsx' | Should -Be 'Studio Rossi'
            Get-HDNomeCliente 'Censimento_Studio-Rossi (2).xlsx' | Should -Be 'Studio Rossi'
            Get-HDNomeCliente 'Censimento.xlsx' | Should -Be ''
        }
    }
}

Describe 'Storico dei cambiamenti' {
    It 'registra valori cambiati, primo censimento e programmi installati o tolti; ignora build e celle vuote' {
        InModuleScope HelpDeskToolkit {
            $nuovaTabella = {
                param($Righe)
                $l = [System.Collections.Generic.List[object]]::new()
                foreach ($r in $Righe) { $l.Add($r) }
                [pscustomobject]@{ Righe = $l }
            }
            $principale = & $nuovaTabella @(
                [pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; OS = 'Microsoft Windows 11 Pro'; OSVersion = '24H2 (26100.2033)'; Memory = '8 GB DDR4'; NomeUtente = 'mario' } }
                [pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'PC-02'; IndirizzoIP = '192.168.1.22' } }
            )
            $tabella = [pscustomobject]@{
                Righe        = $principale.Righe
                Dettagli     = (& $nuovaTabella @([pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; Rilevato = '01/10/2026 09:00'; Office = 'Microsoft 365 Apps for business (16.0.18000)' } }))
                DatiSoftware = (& $nuovaTabella @([pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-01'; SerialNumber = 'S1'; Programmi = "7-Zip 24.08 (x64)`t24.08`tIgor Pavlov`nAdobe Acrobat Reader`t24.1`tAdobe`nTeamViewer`t15.50`tTeamViewer" } }))
            }
            $foto = Get-HDFotoCensimento $tabella

            # nuovo censimento: più RAM, build nuova (non conta), stessa versione di Office a parte il numero
            $pc1 = $tabella.Righe[0].Valori
            $pc1.Memory = '16 GB DDR4 (2 x 8 GB)'
            $pc1.OSVersion = '24H2 (26100.4061)'
            $pc1.Dominio = 'studio.local'
            $tabella.Dettagli.Righe[0].Valori.Rilevato = '05/11/2026 10:00'
            $tabella.Dettagli.Righe[0].Valori.Office = 'Microsoft 365 Apps for business (16.0.19127)'
            $tabella.DatiSoftware.Righe[0].Valori.Programmi = "7-Zip 24.09 (x64)`t24.09`tIgor Pavlov`nTeamViewer`t15.60`tTeamViewer`nAnyDesk`t9.0`tAnyDesk"
            # la riga arrivata dalla scansione ora è censita, e c'è un PC nuovo
            $tabella.Righe[1].Valori.OS = 'Microsoft Windows 10 Pro'
            $tabella.Righe[1].Valori.Manufacturer = 'HP'
            $tabella.Righe[1].Valori.Model = 'HP ProDesk 400 G7'
            $tabella.Righe.Add([pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'PC-03'; OS = 'Microsoft Windows 11 Pro' } })

            $voci = @(Get-HDCambiamenti -Foto $foto -Tabella $tabella -Adesso ([datetime]'2026-11-07 12:00'))
            @($voci | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Data, $_.Hostname, $_.Campo, $_.Prima, $_.Dopo }) | Should -Be @(
                '05/11/2026 10:00|PC-01|Memory|8 GB DDR4|16 GB DDR4 (2 x 8 GB)'
                '07/11/2026 12:00|PC-02|Primo censimento||HP ProDesk 400 G7, Microsoft Windows 10 Pro'
                '07/11/2026 12:00|PC-03|Primo censimento||Microsoft Windows 11 Pro'
                '05/11/2026 10:00|PC-01|Programmi installati||AnyDesk'
                '05/11/2026 10:00|PC-01|Programmi disinstallati|Adobe Acrobat Reader|'
            )

            # in fondo al foglio Storico, una riga per voce
            $storico = ConvertTo-HDTabellaFoglio -Fogli @() -Nome 'Storico' -Colonne $script:HDColonneStorico
            $modifiche = @{}
            Add-HDStoricoATabella -Tabella $storico -Voci $voci -Modifiche $modifiche
            @($modifiche.Keys | Sort-Object) | Should -Be @(2, 3, 4, 5, 6)
            $modifiche[2].Campo | Should -Be 'Memory'
            $modifiche[3].ContainsKey('Prima') | Should -BeFalse
        }
    }
}

Describe 'Pagina del riepilogo per il cliente' {
    It 'punti per priorità con i PC e cosa fare, numeri, cambiamenti, elenco dei PC; testo sempre sicuro' {
        $html = InModuleScope HelpDeskToolkit {
            $pc = @(
                [pscustomobject]@{ Nome = 'PC-01'; Utilizzatore = 'Anna <b>Bianchi</b>'; Valori = @{ OS = 'Microsoft Windows 10 Pro'; OSVersion = '22H2 (19045.1)'; Memory = '8 GB DDR4'; Manufacturer = 'Dell Inc.'; Model = 'OptiPlex 7090' }; Dettagli = @{ Windows11 = 'Sì'; Batteria = 'Nessuna'; Rilevato = '02/11/2026 10:00' } }
                [pscustomobject]@{ Nome = 'PC-02'; Utilizzatore = ''; Valori = @{ OS = 'Microsoft Windows 11 Pro'; OSVersion = '25H2 (26200.1)'; Memory = '16 GB' }; Dettagli = @{ Batteria = '90% (45 di 50 Wh)'; Rilevato = '05/11/2026 10:00'; Office = 'Microsoft 365 Apps for business' } }
            )
            $punti = @(Get-HDPuntiRiepilogo -PC $pc -Adesso ([datetime]'2026-11-07'))
            $cambi = @([pscustomobject]@{ Data = '05/11/2026 10:00'; Hostname = 'PC-02'; Campo = 'Memory'; Prima = '8 GB'; Dopo = '16 GB' }, [pscustomobject]@{ Data = '01/01/2025 10:00'; Hostname = 'PC-01'; Campo = 'Vecchio'; Prima = 'a'; Dopo = 'b' })
            ConvertTo-HDHtmlRiepilogo -Cliente 'Studio Rossi' -PC $pc -Punti $punti -Cambiamenti $cambi -File 'Censimento_Studio_Rossi.xlsx' -Adesso ([datetime]'2026-11-07 15:00')
        }
        # le lettere accentate diventano entità (pi&#249;): si controlla il testo decodificato
        $testo = [System.Net.WebUtility]::HtmlDecode($html)
        $html | Should -Match '<h1>Studio Rossi</h1>'
        $html | Should -Match '1 punto urgente'
        $html | Should -Match '<h2>Da fare subito</h2>'
        $testo | Should -Match 'Windows non riceve più aggiornamenti di sicurezza'
        $testo | Should -Match 'Windows 10, compatibile con Windows 11'
        $html | Should -Match '<b>Cosa fare:</b>'
        $html | Should -Match 'Portatili / fissi</dt><dd>1 / 1'
        $html | Should -Match 'Cambiamenti degli ultimi 90 giorni'
        $html | Should -Not -Match 'Vecchio'
        $html | Should -Match 'Elenco dei PC'
        $html | Should -Match 'Anna &lt;b&gt;Bianchi&lt;/b&gt;'
        $html | Should -Not -Match '<b>Bianchi'
        $html | Should -Not -Match 'tua azienda'
    }
}

Describe 'Riepilogo e Storico nell''Excel' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:CartellaCensimentoTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        Get-ChildItem -LiteralPath $script:CartellaReportTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        $script:File = (New-Censimento -Cliente 'Studio Rossi' -Confirm:$false).FullName
    }

    It 'a ogni censimento il foglio Riepilogo si ricalcola e lo Storico registra i cambiamenti' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        (Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null).Scritto | Should -BeTrue
        $fogli = Get-Fogli $script:File
        $riepilogo = @(Get-RigheFoglio $fogli 'Riepilogo')
        $windows = $riepilogo | Where-Object { $_[2] -eq 'Windows non riceve più aggiornamenti di sicurezza' }
        $windows[1] | Should -Be 'Urgente'
        $windows[3] | Should -Be '1'
        $windows[4] | Should -Be 'PC-AMMINISTRAZIONE (Windows 10, compatibile con Windows 11)'
        $riepilogo[-1][2] | Should -BeLike 'Calcolato il * su 1 PC censito*'
        $storico = @(Get-RigheFoglio $fogli 'Storico')
        $storico.Count | Should -Be 1
        $storico[0][4] | Should -Be 'Primo censimento'

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
        @($storico | ForEach-Object { $_[4] }) | Should -Be @('Primo censimento', 'Memory', 'Accesso remoto', 'Programmi installati')
        ($storico | Where-Object { $_[4] -eq 'Memory' })[6] | Should -BeLike '32 GB*'
        ($storico | Where-Object { $_[4] -eq 'Programmi installati' })[6] | Should -Be 'AnyDesk'
        $remoto = @(Get-RigheFoglio $fogli 'Riepilogo') | Where-Object { $_[2] -eq 'Programmi di accesso remoto' }
        $remoto[4] | Should -Be 'PC-AMMINISTRAZIONE (TeamViewer; AnyDesk)'
    }

    It 'Export-RiepilogoCensimento crea la pagina nella cartella Report' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        $pagina = Export-RiepilogoCensimento -Censimento $script:File -Confirm:$false
        $pagina.DirectoryName | Should -Be $script:CartellaReportTest
        $pagina.Name | Should -BeLike 'Riepilogo_Studio-Rossi_*.html'
        $testo = Get-Content -LiteralPath $pagina.FullName -Raw -Encoding UTF8
        $testo | Should -Match '<h1>Studio Rossi</h1>'
        $testo | Should -Match 'PC-AMMINISTRAZIONE'
        $testo | Should -Match 'Primo censimento'
    }
}

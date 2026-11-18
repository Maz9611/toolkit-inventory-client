BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    . (Join-Path $PSScriptRoot 'DatiDiProva.ps1')

    # Voci di "Installazione applicazioni" come le lascia un PC con AutoCAD, AutoCAD LT e Revit
    function New-ProgrammiAutodesk {
        @(
            [pscustomobject]@{ DisplayName = 'AutoCAD 2024 - English'; DisplayVersion = '24.3.61.0'; Publisher = 'Autodesk, Inc.'; InstallDate = '20240315' }
            [pscustomobject]@{ DisplayName = 'AutoCAD 2024 Language Pack - Italiano'; DisplayVersion = '24.3.61.0'; Publisher = 'Autodesk, Inc.' }
            [pscustomobject]@{ DisplayName = 'AutoCAD LT 2019 - Italiano'; DisplayVersion = '23.0.47.0'; Publisher = 'Autodesk'; InstallDate = '20180620' }
            [pscustomobject]@{ DisplayName = 'Autodesk Revit 2024'; DisplayVersion = '24.2.0.63'; Publisher = 'Autodesk'; InstallDate = '20240110' }
            [pscustomobject]@{ DisplayName = 'Revit Content Libraries 2024'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'Autodesk Material Library 2024'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'Autodesk Genuine Service'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'Autodesk Access'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'Autodesk Desktop Connector'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'DWG TrueView 2024 - English'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'Autodesk Single Sign On Component'; Publisher = 'Autodesk' }
            [pscustomobject]@{ DisplayName = 'AutoCAD Plugin di qualcun altro'; Publisher = 'Altro Editore' }
            [pscustomobject]@{ DisplayName = 'Mozilla Firefox (x64 it)'; Publisher = 'Mozilla' }
        )
    }
    # Uscita di "AdskLicensingInstHelper list": AutoCAD con utente, Revit con la licenza di una Collection
    $script:ListaLicenze = @'
Autodesk Licensing Installer Helper
[
  { "feature_id": "ACD", "def_prod_key": "001P1", "def_prod_ver": "2024.0.0.F", "sel_prod_key": "001P1", "sel_prod_ver": "2024.0.0.F", "sel_mode": "USER", "sel_srv_names": [] },
  { "feature_id": "RVT", "def_prod_key": "829P1", "def_prod_ver": "2024.0.0.F", "sel_prod_key": "02HP1", "sel_prod_ver": "2024.0.0.F", "sel_mode": "USER", "sel_srv_names": null }
]
'@
    $script:LoginState = '<?xml version="1.0" encoding="utf-8"?><LoginState><UserInfo><UserName>mrossi</UserName><Email>Mario.Rossi@Studio-Rossi.it</Email></UserInfo></LoginState>'
    function New-AutodeskGrezzo {
        param([string]$Licenze = $script:ListaLicenze, [object[]]$Avvii = @(
                [pscustomobject]@{ Programma = 'ACAD'; Data = [datetime]'2026-10-01 09:12' }
                [pscustomobject]@{ Programma = 'ACAD'; Data = [datetime]'2025-02-01 09:12' }
                [pscustomobject]@{ Programma = 'REVIT'; Data = [datetime]'2024-01-10 17:40' }
            ), [switch]$SenzaPrefetch)
        $a = [pscustomobject]@{
            Licenze = $Licenze
            Account = @([pscustomobject]@{ Utente = 'mario'; Testo = $script:LoginState })
            Rete    = @()
            Seriali = @([pscustomobject]@{ Prodotto = 'AutoCAD LT 2019'; Seriale = '666-69696969' }, [pscustomobject]@{ Prodotto = 'AutoCAD 2024'; Seriale = '' })
            Avvii   = $Avvii
            Note    = @()
        }
        if ($SenzaPrefetch) { $a.Avvii = $null }
        $a
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Prodotti Autodesk: versione, licenza, account, ultimo avvio' {
    It 'tiene solo i prodotti veri (niente componenti, librerie, servizi, visualizzatori)' {
        $p = @(InModuleScope HelpDeskToolkit -Parameters @{ P = (New-ProgrammiAutodesk) } { param($P) Select-HDProdottiAutodesk $P })
        @($p | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Nome, $_.Famiglia, $_.Anno, $_.Installato }) | Should -Be @(
            'AutoCAD 2024|AutoCAD|2024|15/03/2024'
            'AutoCAD LT 2019|AutoCAD LT|2019|20/06/2018'
            'Revit 2024|Revit|2024|10/01/2024'
        )
    }

    It 'legge la licenza registrata, anche quando arriva da una Collection' {
        $l = @(InModuleScope HelpDeskToolkit -Parameters @{ T = $script:ListaLicenze } { param($T) ConvertFrom-HDLicenzeAutodesk $T })
        $l.Count | Should -Be 2
        $l[0].Chiave | Should -Be '001P1'
        $l[0].Anno | Should -Be '2024'
        $l[0].Modalita | Should -Be 'USER'
        $l[1].ChiaveScelta | Should -Be '02HP1'
        @(InModuleScope HelpDeskToolkit { ConvertFrom-HDLicenzeAutodesk 'niente di leggibile' }) | Should -BeNullOrEmpty
    }

    It 'una riga per prodotto con licenza in parole, account, ultimo avvio, seriale' {
        $righe = InModuleScope HelpDeskToolkit -Parameters @{ P = (New-ProgrammiAutodesk); A = (New-AutodeskGrezzo) } { param($P, $A) ConvertTo-HDAutodeskPC -Programmi $P -Autodesk $A }
        @($righe).Count | Should -Be 3
        $acad = $righe | Where-Object Prodotto -eq 'AutoCAD 2024'
        $acad.Licenza | Should -Be 'Utente con nome (abbonamento, account Autodesk)'
        $acad.Account | Should -Be 'mario.rossi@studio-rossi.it (mario)'
        $acad.UltimoAvvio | Should -Be '01/10/2026'
        $acad.Versione | Should -Be '24.3.61.0'
        $acad.Seriale | Should -Be ''
        $lt = $righe | Where-Object Prodotto -eq 'AutoCAD LT 2019'
        $lt.Licenza | Should -Be 'Singola con seriale (versione 2019 o precedente)'
        $lt.Seriale | Should -Be '666-69696969 (seriale delle copie non originali)'
        $lt.UltimoAvvio | Should -Be 'Mai (dall''ultima pulizia di Windows)'
        $revit = $righe | Where-Object Prodotto -eq 'Revit 2024'
        $revit.Licenza | Should -Be 'Utente con nome (abbonamento, account Autodesk), con la licenza di un''altra offerta (chiave 02HP1, es. una Collection)'
        $revit.UltimoAvvio | Should -Be '10/01/2024'
    }

    It 'licenza di rete, nessuna licenza scelta, versioni vecchie in rete, ultimo avvio senza diritti' {
        InModuleScope HelpDeskToolkit {
            $rete = [pscustomobject]@{ Chiave = '001P1'; Anno = '2024'; ChiaveScelta = '001P1'; Modalita = 'NETWORK'; Server = 'licsrv01' }
            Format-HDLicenzaAutodesk -Licenza $rete -Anno '2024' | Should -Be 'Di rete (server licsrv01)'
            $nessuna = [pscustomobject]@{ Chiave = '001P1'; Anno = '2024'; ChiaveScelta = ''; Modalita = ''; Server = '' }
            Format-HDLicenzaAutodesk -Licenza $nessuna -Anno '2024' | Should -Be 'Nessuna licenza scelta'
            Format-HDLicenzaAutodesk -Licenza $null -Anno '2018' -Rete 'licsrv01' | Should -Be 'Di rete FlexNet (server licsrv01)'
            Format-HDLicenzaAutodesk -Licenza $null -Anno '2024' | Should -Be ''
        }
        $righe = InModuleScope HelpDeskToolkit -Parameters @{ P = (New-ProgrammiAutodesk); A = (New-AutodeskGrezzo -SenzaPrefetch -Licenze '') } { param($P, $A) ConvertTo-HDAutodeskPC -Programmi $P -Autodesk $A }
        ($righe | Where-Object Prodotto -eq 'AutoCAD 2024').UltimoAvvio | Should -Be 'Serve avviare come amministratore'
        ($righe | Where-Object Prodotto -eq 'AutoCAD 2024').Licenza | Should -Be 'Non letta'
        # niente dati Autodesk raccolti: $null; raccolti ma nessun prodotto: elenco vuoto
        InModuleScope HelpDeskToolkit { $null -eq (ConvertTo-HDAutodeskPC -Programmi @() -Autodesk $null) } | Should -BeTrue
        InModuleScope HelpDeskToolkit -Parameters @{ A = (New-AutodeskGrezzo) } { param($A) $x = ConvertTo-HDAutodeskPC -Programmi @() -Autodesk $A; ($null -ne $x) -and (@($x).Count -eq 0) } | Should -BeTrue
    }

    It 'foglio Autodesk riscritto: i PC appena censiti si sostituiscono, gli altri restano, le Note scritte a mano pure' {
        InModuleScope HelpDeskToolkit {
            $esistenti = @(
                [pscustomobject]@{ Numero = 2; Valori = @{ Hostname = 'PC-A'; SerialNumber = 'SA'; Prodotto = 'AutoCAD 2024'; Licenza = 'Utente con nome'; Note = 'licenza di Mario' } }
                [pscustomobject]@{ Numero = 3; Valori = @{ Hostname = 'PC-A'; SerialNumber = 'SA'; Prodotto = 'AutoCAD LT 2019'; Note = 'da disinstallare' } }
                [pscustomobject]@{ Numero = 4; Valori = @{ Hostname = 'PC-B'; SerialNumber = 'SB'; Prodotto = 'Revit 2024' } }
            )
            $nuovo = [pscustomobject]@{ Hostname = 'PC-A'; SerialNumber = 'SA'; Utilizzatore = 'Mario Rossi'; Rilevato = '18/11/2026 10:00'; Righe = @([pscustomobject]@{ Prodotto = 'AutoCAD 2024'; Licenza = 'Di rete (server licsrv01)'; Versione = '24.3' }) }
            $righe = @(Get-HDRigheAutodesk -Esistenti $esistenti -Nuovi @($nuovo))
            @($righe | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_['Hostname'], $_['Prodotto'], $_['Licenza'], $_['Note'] }) | Should -Be @(
                'PC-A|AutoCAD 2024|Di rete (server licsrv01)|licenza di Mario'
                'PC-B|Revit 2024||'
            )
            $righe[0]['Utilizzatore'] | Should -Be 'Mario Rossi'
        }
    }

    It 'nel riepilogo: seriale delle copie non originali, prodotto senza licenza, prodotto fermo da mesi' {
        $pc = [pscustomobject]@{
            Nome = 'PC-CAD'; Utilizzatore = ''; Valori = @{ Hostname = 'PC-CAD'; OS = 'Microsoft Windows 11 Pro'; OSVersion = '25H2' }; Dettagli = @{ Rilevato = '18/11/2026 10:00' }
            Autodesk = @(
                @{ Prodotto = 'AutoCAD LT 2019'; Licenza = 'Singola con seriale (versione 2019 o precedente)'; Seriale = '666-69696969 (seriale delle copie non originali)'; UltimoAvvio = '10/01/2024' }
                @{ Prodotto = 'AutoCAD 2024'; Licenza = 'Nessuna licenza scelta'; UltimoAvvio = '01/11/2026' }
            )
        }
        $punti = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $pc } { param($P) Get-HDPuntiRiepilogo -PC @($P) -Adesso ([datetime]'2026-11-18') })
        ($punti | Where-Object Punto -eq 'Autodesk con seriale delle copie non originali').PC[0].Fatto | Should -Be 'AutoCAD LT 2019: seriale delle copie non originali (666-69696969)'
        ($punti | Where-Object Punto -eq 'Autodesk con seriale delle copie non originali').Priorita | Should -Be 'Urgente'
        ($punti | Where-Object Punto -eq 'Autodesk senza licenza').PC[0].Fatto | Should -Be 'AutoCAD 2024 senza licenza'
        ($punti | Where-Object Punto -eq 'Autodesk non usato da tempo').PC[0].Fatto | Should -Be 'AutoCAD LT 2019 non avviato dal 10/01/2024'
    }
}

Describe 'Foglio Autodesk nell''Excel' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:CartellaCensimentoTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        Get-ChildItem -LiteralPath $script:CartellaReportTest -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force
        $script:File = (New-Censimento -Cliente 'Studio CAD' -Confirm:$false).FullName
    }

    It 'nasce solo quando un PC ha prodotti Autodesk; un nuovo censimento aggiorna le righe di quel PC' {
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario { New-GrezzoCompleto }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        @(Get-Fogli $script:File | ForEach-Object { $_.Nome }) | Should -Not -Contain 'Autodesk'

        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario {
            $g = New-GrezzoCompleto -Nome 'PC-CAD' -Serie 'SER-CAD' -Mac '00:11:22:33:44:CA' -Programmi (New-ProgrammiAutodesk)
            $g | Add-Member -NotePropertyName Autodesk -NotePropertyValue (New-AutodeskGrezzo)
            $g
        }
        (Add-PCCensimento -Censimento $script:File -NonChiedere -Utilizzatore 'Ufficio tecnico' -Confirm:$false 6>$null).Scritto | Should -BeTrue
        $fogli = Get-Fogli $script:File
        $righe = @(Get-RigheFoglio $fogli 'Autodesk')
        @($righe | ForEach-Object { '{0}|{1}|{2}' -f $_[1], $_[2], $_[3] }) | Should -Be @('PC-CAD|Ufficio tecnico|AutoCAD 2024', 'PC-CAD|Ufficio tecnico|AutoCAD LT 2019', 'PC-CAD|Ufficio tecnico|Revit 2024')
        $righe[0][6] | Should -Be 'mario.rossi@studio-rossi.it (mario)'
        $riepilogo = @(Get-RigheFoglio $fogli 'Riepilogo') | Where-Object { $_[1] -eq 'PC-CAD' }
        $riepilogo[3] | Should -Match 'AutoCAD LT 2019: seriale delle copie non originali'

        # AutoCAD LT disinstallato: al censimento dopo la sua riga sparisce
        Mock -ModuleName HelpDeskToolkit Invoke-HDRaccoltaInventario {
            $g = New-GrezzoCompleto -Nome 'PC-CAD' -Serie 'SER-CAD' -Mac '00:11:22:33:44:CA' -Programmi @(New-ProgrammiAutodesk | Where-Object { $_.DisplayName -notlike 'AutoCAD LT*' })
            $g | Add-Member -NotePropertyName Autodesk -NotePropertyValue (New-AutodeskGrezzo)
            $g
        }
        [void](Add-PCCensimento -Censimento $script:File -NonChiedere -Confirm:$false 6>$null)
        $fogli = Get-Fogli $script:File
        @(Get-RigheFoglio $fogli 'Autodesk' | ForEach-Object { $_[3] }) | Should -Be @('AutoCAD 2024', 'Revit 2024')
        @(Get-RigheFoglio $fogli 'Storico' | ForEach-Object { $_[4] }) | Should -Contain 'Programmi disinstallati: AutoCAD LT 2019 - Italiano'

        # e la pagina per il cliente ha la sezione Autodesk
        $pagina = Export-RiepilogoCensimento -Censimento $script:File -Confirm:$false
        $testo = [System.Net.WebUtility]::HtmlDecode((Get-Content -LiteralPath $pagina.FullName -Raw -Encoding UTF8))
        $testo | Should -Match '<h2>Software Autodesk</h2>'
        $testo | Should -Match 'Revit 2024</td><td>Utente con nome'
        $testo | Should -Match 'Account Autodesk con cui si è entrati sui PC: mario.rossi@studio-rossi.it'
    }
}

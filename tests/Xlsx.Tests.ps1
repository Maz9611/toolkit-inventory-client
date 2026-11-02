BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    Add-Type -AssemblyName System.IO.Compression

    # Crea un .xlsx a partire dai testi XML delle sue parti
    function New-XlsxGrezzo([string]$Percorso, [System.Collections.Specialized.OrderedDictionary]$Parti) {
        $flusso = [System.IO.File]::Open($Percorso, 'Create')
        $zip = New-Object System.IO.Compression.ZipArchive($flusso, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            foreach ($n in $Parti.Keys) {
                $s = $zip.CreateEntry($n).Open()
                $b = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Parti[$n])
                $s.Write($b, 0, $b.Length)
                $s.Dispose()
            }
        }
        finally { $zip.Dispose() }
    }
    function Get-VoceZip([string]$Percorso, [string]$Nome) {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Percorso)
        try { $r = New-Object System.IO.StreamReader($zip.GetEntry($Nome).Open()); $r.ReadToEnd() } finally { $zip.Dispose() }
    }

    $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $script:PartiComeExcel = [ordered]@{
        '[Content_Types].xml'                 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'
        '_rels/.rels'                         = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'                     = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook $ns><sheets><sheet name=`"Info`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"Elenco PC`" sheetId=`"2`" r:id=`"rId2`"/></sheets></workbook>"
        'xl/_rels/workbook.xml.rels'          = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
        'xl/styles.xml'                       = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFEFEF"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>'
        'xl/sharedStrings.xml'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6"><si><t>Hostname</t></si><si><t>Indirizzo IP</t></si><si><t>Cynet</t></si><si><t>PC-01</t></si><si><r><t>PC-</t></r><r><rPr><b/></rPr><t>02</t></r><rPh sb="0" eb="1"><t>X</t></rPh></si><si><t>Elenco dei PC</t></si></sst>'
        'xl/worksheets/sheet1.xml'            = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet $ns><sheetData><row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>Solo note</t></is></c></row></sheetData></worksheet>"
        'xl/worksheets/sheet2.xml'            = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><worksheet $ns xmlns:mc=`"http://schemas.openxmlformats.org/markup-compatibility/2006`" mc:Ignorable=`"x14ac`" xmlns:x14ac=`"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac`"><dimension ref=`"A1:C4`"/><sheetData><row r=`"1`"><c r=`"A1`" t=`"s`"><v>5</v></c></row><row r=`"2`" spans=`"1:3`"><c r=`"A2`" t=`"s`"><v>0</v></c><c r=`"B2`" t=`"s`"><v>1</v></c><c r=`"C2`" t=`"s`"><v>2</v></c></row><row r=`"3`" spans=`"1:3`" x14ac:dyDescent=`"0.25`"><c r=`"A3`" s=`"2`" t=`"s`"><v>3</v></c><c r=`"B3`" s=`"2`"><v>10</v></c></row><row r=`"4`" spans=`"1:3`"><c r=`"A4`" s=`"2`" t=`"s`"><v>4</v></c><c r=`"C4`" s=`"2`"><f>1+1</f><v>2</v></c></row></sheetData><tableParts count=`"1`"><tablePart r:id=`"rId1`"/></tableParts></worksheet>"
        'xl/worksheets/_rels/sheet2.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'
        'xl/tables/table1.xml'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tabella1" displayName="Tabella1" ref="A2:C4"><autoFilter ref="A2:C4"/><tableColumns count="3"><tableColumn id="1" name="Hostname"/><tableColumn id="2" name="Indirizzo IP"/><tableColumn id="3" name="Cynet"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>'
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Riferimenti di cella' {
    It 'colonna <Indice> = <Lettere>' -TestCases @(
        @{ Indice = 1; Lettere = 'A' }
        @{ Indice = 26; Lettere = 'Z' }
        @{ Indice = 27; Lettere = 'AA' }
        @{ Indice = 703; Lettere = 'AAA' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ I = $Indice } { param($I) ConvertTo-HDLettereColonna $I } | Should -Be $Lettere
        (InModuleScope HelpDeskToolkit -Parameters @{ L = $Lettere } { param($L) ConvertFrom-HDRiferimento "$($L)7" }).Colonna | Should -Be $Indice
    }

    It 'allarga un intervallo solo se finiva subito sopra le righe nuove' {
        InModuleScope HelpDeskToolkit {
            Expand-HDIntervallo -Intervallo 'A1:T10' -PrimaNuova 11 -UltimaRiga 15 | Should -Be 'A1:T15'
            Expand-HDIntervallo -Intervallo 'A1:T20' -PrimaNuova 11 -UltimaRiga 15 | Should -BeNullOrEmpty
            Expand-HDIntervallo -Intervallo 'A1:T5' -PrimaNuova 11 -UltimaRiga 15 | Should -BeNullOrEmpty
        }
    }
}

Describe 'File creato dal toolkit' {
    BeforeAll {
        $script:Nuovo = Join-Path $TestDrive 'nuovo.xlsx'
        InModuleScope HelpDeskToolkit -Parameters @{ P = $script:Nuovo } {
            param($P)
            $col = @(
                @{ Titolo = 'Hostname'; Larghezza = 18; Testo = $true }
                @{ Titolo = 'Indirizzo IP'; Larghezza = 15; Testo = $true }
                @{ Titolo = 'Free Disk Space %'; Larghezza = 10 }
                @{ Titolo = 'Locazione'; Larghezza = 18; Manuale = $true; Testo = $true }
            )
            New-HDXlsx -Percorso $P -NomeFoglio 'Censimento' -Colonne $col -Elenchi @{ 2 = 'DHCP,Statico' } -Evidenzia @{ 3 = 'AND(ISNUMBER({0}),{0}<10)' }
        }
    }

    It 'ha l''intestazione giusta, bloccata e con il filtro' {
        $f = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:Nuovo } { param($P) Read-HDXlsx $P })
        $f[0].Nome | Should -Be 'Censimento'
        @($f[0].Righe[1].Values) | Should -Contain 'Free Disk Space %'
        $xml = Get-VoceZip $script:Nuovo 'xl/worksheets/sheet1.xml'
        $xml | Should -Match 'state="frozen"'
        $xml | Should -Match '<autoFilter ref="A1:D1"/>'
        $xml | Should -Match 'AND\(ISNUMBER\(C2\),C2&lt;10\)'
        $xml | Should -Match '<c r="D1" s="2"'
    }

    It 'aggiunge e aggiorna righe: numeri come numeri, testo protetto, stile della colonna' {
        InModuleScope HelpDeskToolkit -Parameters @{ P = $script:Nuovo } {
            param($P)
            [void](Update-HDXlsx -Percorso $P -ParteFoglio 'xl/worksheets/sheet1.xml' -IndiceFoglio 0 -ColonneNumeriche @(3) -Modifiche @(
                    @{ Riga = 2; Celle = @{ 4 = 'Reception'; 1 = 'PC-01 & <x>'; 3 = '25' } }
                    @{ Riga = 3; Celle = @{ 1 = 'PC-02' } }))
            [void](Update-HDXlsx -Percorso $P -ParteFoglio 'xl/worksheets/sheet1.xml' -IndiceFoglio 0 -Modifiche @(@{ Riga = 2; Celle = @{ 2 = '192.168.1.10' } }))
        }
        $f = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:Nuovo } { param($P) Read-HDXlsx $P })
        $f[0].Righe[2][1] | Should -Be 'PC-01 & <x>'
        $f[0].Righe[2][2] | Should -Be '192.168.1.10'
        $f[0].Righe[2][4] | Should -Be 'Reception'
        $xml = Get-VoceZip $script:Nuovo 'xl/worksheets/sheet1.xml'
        $xml | Should -Match '<c r="C2" s="0"><v>25</v></c>'
        $xml | Should -Match '<row r="2"><c r="A2" s="3" t="inlineStr"><is><t>PC-01 &amp; &lt;x&gt;</t></is></c><c r="B2" s="3"'
        $xml | Should -Match '<autoFilter ref="A1:D3"'
        Get-VoceZip $script:Nuovo 'xl/workbook.xml' | Should -Match '\$A\$1:\$D\$3'
    }
}

Describe 'File salvato da Excel (stringhe condivise, tabella, formule)' {
    BeforeEach {
        $script:ComeExcel = Join-Path $TestDrive 'come-excel.xlsx'
        Remove-Item -LiteralPath $script:ComeExcel -ErrorAction SilentlyContinue
        New-XlsxGrezzo $script:ComeExcel $script:PartiComeExcel
    }

    It 'aggiunge un foglio senza toccare gli altri (relazioni, tipi, stili, filtro)' {
        $nuovo = InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } {
            param($P)
            $c = Add-HDFoglioXlsx -Percorso $P -NomeFoglio 'Software' -Colonne $script:HDColonneSoftware
            [void](Update-HDXlsx -Percorso $P -ParteFoglio $c.Parte -IndiceFoglio $c.Indice -Modifiche @(@{ Riga = 2; Celle = @{ 1 = 'PC-01'; 2 = 'Firefox' } }))
            $c
        }
        $nuovo.Parte | Should -Be 'xl/worksheets/sheet3.xml'
        $nuovo.Indice | Should -Be 2
        $f = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } { param($P) Read-HDXlsx $P })
        ($f | ForEach-Object { $_.Nome }) | Should -Be @('Info', 'Elenco PC', 'Software')
        $f[2].Righe[1][2] | Should -Be 'Programma'
        $f[2].Righe[2][2] | Should -Be 'Firefox'
        $f[1].Righe[4][1] | Should -Be 'PC-02'
        Get-VoceZip $script:ComeExcel '[Content_Types].xml' | Should -Match 'PartName="/xl/worksheets/sheet3.xml"'
        $rels = Get-VoceZip $script:ComeExcel 'xl/_rels/workbook.xml.rels'
        ([regex]::Matches($rels, '<Relationship ')).Count | Should -Be 5
        $wb = Get-VoceZip $script:ComeExcel 'xl/workbook.xml'
        $wb | Should -Match 'name="Software" sheetId="3" r:id="rIdHD1"'
        $wb | Should -Match "_xlnm._FilterDatabase`" localSheetId=`"2`" hidden=`"1`">'Software'!\`$A\`$1:\`$G\`$2"
        $stili = Get-VoceZip $script:ComeExcel 'xl/styles.xml'
        $stili | Should -Match '<fonts count="2">'
        $stili | Should -Match '<cellXfs count="5">'
        Get-VoceZip $script:ComeExcel 'xl/worksheets/sheet3.xml' | Should -Match '<c r="A1" s="3" t="inlineStr">'
        { InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } { param($P) Add-HDFoglioXlsx -Percorso $P -NomeFoglio 'Software' -Colonne $script:HDColonneSoftware } } | Should -Throw '*esiste già*'
    }

    It 'legge fogli, stringhe condivise e testo formattato a pezzi' {
        $f = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } { param($P) Read-HDXlsx $P })
        ($f | ForEach-Object { $_.Nome }) | Should -Be @('Info', 'Elenco PC')
        $f[1].Parte | Should -Be 'xl/worksheets/sheet2.xml'
        $f[1].Righe[2][1] | Should -Be 'Hostname'
        $f[1].Righe[4][1] | Should -Be 'PC-02'
        $f[1].Righe[3][2] | Should -Be '10'
    }

    It 'allarga la tabella, copia lo stile della riga sopra e non tocca le formule' {
        $saltate = InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } {
            param($P)
            Update-HDXlsx -Percorso $P -ParteFoglio 'xl/worksheets/sheet2.xml' -IndiceFoglio 1 -RigaIntestazione 2 -Modifiche @(
                @{ Riga = 4; Celle = @{ 2 = '192.168.1.5'; 3 = 'Sì' } }
                @{ Riga = 5; Celle = @{ 1 = 'PC-03'; 3 = 'No' } })
        }
        @($saltate) | Should -Be @('C4')
        $tab = Get-VoceZip $script:ComeExcel 'xl/tables/table1.xml'
        $tab | Should -Match 'ref="A2:C5"'
        $tab | Should -Match '<autoFilter ref="A2:C5"'
        $xml = Get-VoceZip $script:ComeExcel 'xl/worksheets/sheet2.xml'
        $xml | Should -Match '<c r="A5" s="2" t="inlineStr"><is><t>PC-03</t></is></c>'
        $xml | Should -Match '<c r="C4" s="2"><f>1\+1</f>'
        $xml | Should -Match 'mc:Ignorable="x14ac"'
        $xml | Should -Match '<row r="4"><c r="A4"'
        $f = @(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } { param($P) Read-HDXlsx $P })
        $f[1].Righe[4][2] | Should -Be '192.168.1.5'
        $f[1].Righe[4][1] | Should -Be 'PC-02'
        $f[0].Righe[1][1] | Should -Be 'Solo note'
    }

    It 'con un foglio che usa un prefisso (x:) crea le celle con lo stesso prefisso' {
        $p = Join-Path $TestDrive 'prefisso.xlsx'
        $parti = [ordered]@{}
        foreach ($k in $script:PartiComeExcel.Keys) { $parti[$k] = $script:PartiComeExcel[$k] }
        $parti['xl/worksheets/sheet1.xml'] = '<?xml version="1.0" encoding="UTF-8"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>Hostname</x:t></x:is></x:c></x:row></x:sheetData></x:worksheet>'
        New-XlsxGrezzo $p $parti
        InModuleScope HelpDeskToolkit -Parameters @{ P = $p } { param($P) [void](Update-HDXlsx -Percorso $P -ParteFoglio 'xl/worksheets/sheet1.xml' -IndiceFoglio 0 -Modifiche @(@{ Riga = 2; Celle = @{ 1 = 'PC-09' } })) }
        $xml = Get-VoceZip $p 'xl/worksheets/sheet1.xml'
        $xml | Should -Match '<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t>PC-09</x:t></x:is></x:c></x:row>'
        $xml | Should -Not -Match 'xmlns=""'
    }

    It 'toglie i caratteri che renderebbero il file illeggibile' {
        InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } {
            param($P)
            [void](Update-HDXlsx -Percorso $P -ParteFoglio 'xl/worksheets/sheet2.xml' -IndiceFoglio 1 -Modifiche @(@{ Riga = 6; Celle = @{ 1 = "SN`0123`b" } }))
        }
        (@(InModuleScope HelpDeskToolkit -Parameters @{ P = $script:ComeExcel } { param($P) Read-HDXlsx $P }))[1].Righe[6][1] | Should -Be 'SN123'
    }
}

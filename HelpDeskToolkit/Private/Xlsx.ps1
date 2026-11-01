# Lettura e scrittura di file Excel (.xlsx) senza Excel e senza librerie esterne.
# Un .xlsx è uno zip di file XML (Office Open XML). Le modifiche toccano solo le celle
# indicate: formattazione, filtri, tabelle, convalide e altri fogli restano come sono.

$script:HDNsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:HDNsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:HDNsPkg = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Initialize-HDZip {
    if (-not ('System.IO.Compression.ZipArchive' -as [type])) { Add-Type -AssemblyName System.IO.Compression }
}

function ConvertTo-HDLettereColonna {
    # 1 -> A, 27 -> AA
    param([Parameter(Mandatory)][int]$Indice)
    $s = ''
    while ($Indice -gt 0) {
        $resto = ($Indice - 1) % 26
        $s = [string][char](65 + $resto) + $s
        $Indice = [int][math]::Floor(($Indice - 1) / 26)
    }
    $s
}

function ConvertFrom-HDRiferimento {
    # "AB12" -> @{ Colonna = 28; Riga = 12 }
    param([string]$Riferimento)
    if ($Riferimento -notmatch '^\$?([A-Za-z]{1,3})\$?(\d+)$') { return $null }
    $col = 0
    foreach ($ch in $Matches[1].ToUpperInvariant().ToCharArray()) { $col = $col * 26 + ([int]$ch - 64) }
    @{ Colonna = $col; Riga = [int]$Matches[2] }
}

function Open-HDZip {
    # In lettura il file si apre anche se è già aperto in Excel
    param([Parameter(Mandatory)][string]$Percorso, [switch]$Scrittura)
    Initialize-HDZip
    if ($Scrittura) {
        $flusso = [System.IO.File]::Open($Percorso, 'Open', 'ReadWrite', 'None')
        return New-Object System.IO.Compression.ZipArchive($flusso, [System.IO.Compression.ZipArchiveMode]::Update, $false)
    }
    $flusso = [System.IO.File]::Open($Percorso, 'Open', 'Read', 'ReadWrite')
    New-Object System.IO.Compression.ZipArchive($flusso, [System.IO.Compression.ZipArchiveMode]::Read, $false)
}

function Read-HDVoceZip {
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$Nome)
    $voce = $Zip.GetEntry($Nome)
    if (-not $voce) { return $null }
    $flusso = $voce.Open()
    try {
        $lettore = New-Object System.IO.StreamReader($flusso, [System.Text.Encoding]::UTF8, $true)
        $lettore.ReadToEnd()
    }
    finally { $flusso.Dispose() }
}

function Write-HDVoceZip {
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$Nome, [Parameter(Mandatory)][string]$Testo)
    if ($Zip.Mode -ne [System.IO.Compression.ZipArchiveMode]::Create) {
        $vecchia = $Zip.GetEntry($Nome)
        if ($vecchia) { $vecchia.Delete() }
    }
    $voce = $Zip.CreateEntry($Nome, [System.IO.Compression.CompressionLevel]::Optimal)
    $flusso = $voce.Open()
    try {
        $byte = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Testo)
        $flusso.Write($byte, 0, $byte.Length)
    }
    finally { $flusso.Dispose() }
}

function Read-HDXml {
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$Nome)
    $testo = Read-HDVoceZip -Zip $Zip -Nome $Nome
    if ($null -eq $testo) { return $null }
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($testo)
    , $doc
}

function New-HDNsm {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Doc)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $nsm.AddNamespace('x', $script:HDNsMain)
    $nsm.AddNamespace('r', $script:HDNsRel)
    $nsm.AddNamespace('p', $script:HDNsPkg)
    , $nsm
}

function Resolve-HDParte {
    # Percorso di una parte a partire da quella che la cita: ("xl/workbook.xml", "worksheets/sheet1.xml") -> "xl/worksheets/sheet1.xml"
    param([string]$Origine, [string]$Destinazione)
    if ($Destinazione.StartsWith('/')) { return $Destinazione.TrimStart('/') }
    $cartella = $Origine -replace '[^/]+$', ''
    $parti = New-Object System.Collections.Generic.List[string]
    foreach ($p in (($cartella + $Destinazione) -split '/')) {
        if ($p -eq '..') { if ($parti.Count -gt 0) { $parti.RemoveAt($parti.Count - 1) } }
        elseif ($p -and $p -ne '.') { $parti.Add($p) }
    }
    $parti -join '/'
}

function Get-HDParteRelazione {
    param([string]$Parte)
    $cartella = $Parte -replace '[^/]+$', ''
    $nome = $Parte -replace '^.*/', ''
    '{0}_rels/{1}.rels' -f $cartella, $nome
}

function Get-HDRelazioni {
    # Relazioni (Id, Type, Target risolto) di una parte
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][AllowEmptyString()][string]$Parte)
    $doc = Read-HDXml -Zip $Zip -Nome (Get-HDParteRelazione $Parte)
    if (-not $doc) { return @() }
    $nsm = New-HDNsm $doc
    foreach ($r in $doc.SelectNodes('/p:Relationships/p:Relationship', $nsm)) {
        [pscustomobject]@{
            Id     = $r.GetAttribute('Id')
            Tipo   = $r.GetAttribute('Type')
            Parte  = (Resolve-HDParte -Origine $Parte -Destinazione $r.GetAttribute('Target'))
            Esterna = ($r.GetAttribute('TargetMode') -eq 'External')
        }
    }
}

function Get-HDTestoNodo {
    # Testo di una stringa condivisa o in linea (anche formattata a pezzi), senza la fonetica
    param($Nodo, $Nsm)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $Nodo.SelectNodes('.//x:t[not(ancestor::x:rPh)]', $Nsm)) { [void]$sb.Append($t.InnerText) }
    $sb.ToString()
}

function Get-HDValoreCella {
    param($Cella, $Nsm, $Condivise)
    $v = $Cella.SelectSingleNode('x:v', $Nsm)
    switch ($Cella.GetAttribute('t')) {
        's' {
            if ($v -and $v.InnerText -match '^\d+$' -and [int]$v.InnerText -lt $Condivise.Count) { return $Condivise[[int]$v.InnerText] }
            return ''
        }
        'inlineStr' {
            $is = $Cella.SelectSingleNode('x:is', $Nsm)
            if ($is) { return Get-HDTestoNodo $is $Nsm }
            return ''
        }
        default { if ($v) { return $v.InnerText } else { return '' } }
    }
}

function Get-HDInfoWorkbook {
    # Percorso del workbook, fogli (nome, parte, indice) e stringhe condivise
    param([Parameter(Mandatory)]$Zip)
    $parteWb = 'xl/workbook.xml'
    foreach ($r in @(Get-HDRelazioni -Zip $Zip -Parte '')) {
        if ($r.Tipo -like '*/officeDocument') { $parteWb = $r.Parte }
    }
    $wb = Read-HDXml -Zip $Zip -Nome $parteWb
    if (-not $wb) { throw 'Il file non è un file Excel (.xlsx) valido.' }
    $nsm = New-HDNsm $wb
    $relazioni = @(Get-HDRelazioni -Zip $Zip -Parte $parteWb)
    $fogli = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($s in $wb.SelectNodes('/x:workbook/x:sheets/x:sheet', $nsm)) {
        $id = $s.GetAttribute('id', $script:HDNsRel)
        $rel = $relazioni | Where-Object { $_.Id -eq $id } | Select-Object -First 1
        if ($rel -and $rel.Tipo -like '*/worksheet') {
            $fogli.Add([pscustomobject]@{ Nome = $s.GetAttribute('name'); Parte = $rel.Parte; Indice = $i })
        }
        $i++
    }
    $condivise = New-Object System.Collections.Generic.List[string]
    $relSst = $relazioni | Where-Object { $_.Tipo -like '*/sharedStrings' } | Select-Object -First 1
    if ($relSst) {
        $sst = Read-HDXml -Zip $Zip -Nome $relSst.Parte
        if ($sst) {
            $nsmSst = New-HDNsm $sst
            foreach ($si in $sst.SelectNodes('/x:sst/x:si', $nsmSst)) { $condivise.Add((Get-HDTestoNodo $si $nsmSst)) }
        }
    }
    [pscustomobject]@{ ParteWorkbook = $parteWb; Fogli = $fogli.ToArray(); Condivise = $condivise }
}

function Read-HDXlsx {
    <#
    Legge tutti i fogli: per ognuno Nome, Parte, Indice e Righe (numero riga -> hashtable colonna -> testo).
    #>
    param([Parameter(Mandatory)][string]$Percorso)
    $zip = Open-HDZip -Percorso (Resolve-HDPercorso $Percorso)
    try {
        $info = Get-HDInfoWorkbook -Zip $zip
        foreach ($f in $info.Fogli) {
            $doc = Read-HDXml -Zip $zip -Nome $f.Parte
            $righe = New-Object 'System.Collections.Generic.SortedDictionary[int,hashtable]'
            if ($doc) {
                $nsm = New-HDNsm $doc
                $numero = 0
                foreach ($row in $doc.SelectNodes('/x:worksheet/x:sheetData/x:row', $nsm)) {
                    $r = $row.GetAttribute('r')
                    if ($r) { $numero = [int]$r } else { $numero++ }
                    $celle = @{}
                    $colonna = 0
                    foreach ($c in $row.SelectNodes('x:c', $nsm)) {
                        $rif = ConvertFrom-HDRiferimento $c.GetAttribute('r')
                        if ($rif) { $colonna = $rif.Colonna } else { $colonna++ }
                        $valore = Get-HDValoreCella -Cella $c -Nsm $nsm -Condivise $info.Condivise
                        if ($valore -ne '') { $celle[$colonna] = $valore }
                    }
                    if ($celle.Count -gt 0) { $righe[$numero] = $celle }
                }
            }
            [pscustomobject]@{ Nome = $f.Nome; Parte = $f.Parte; Indice = $f.Indice; Righe = $righe }
        }
    }
    finally { $zip.Dispose() }
}

function New-HDElementoXml {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param([System.Xml.XmlDocument]$Doc, [string]$Nome)
    , $Doc.CreateElement($Doc.DocumentElement.Prefix, $Nome, $script:HDNsMain)
}

function ConvertTo-HDTestoXmlSicuro {
    # Toglie i caratteri di controllo che renderebbero il file illeggibile
    param([string]$Testo)
    [regex]::Replace([string]$Testo, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
}

function Expand-HDIntervallo {
    # "A1:T10" -> "A1:T25" se l'intervallo finiva subito sopra le righe aggiunte
    param([string]$Intervallo, [int]$PrimaNuova, [int]$UltimaRiga)
    if ($Intervallo -notmatch '^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$') { return $null }
    $inizio = [int]$Matches[2]
    $fine = [int]$Matches[4]
    if ($fine -lt $UltimaRiga -and $fine -ge $PrimaNuova - 1 -and $inizio -lt $PrimaNuova) {
        return '{0}{1}:{2}{3}' -f $Matches[1], $inizio, $Matches[3], $UltimaRiga
    }
    $null
}

function Update-HDXlsx {
    <#
    Scrive celle in un foglio esistente. $Modifiche: elenco di @{ Riga = 5; Celle = @{ 3 = 'testo'; 17 = 25 } }.
    Le celle con una formula non vengono toccate. Le righe nuove prendono lo stile della
    riga sopra (o della colonna); tabelle, filtro automatico e dimensione si allargano.
    Restituisce i riferimenti delle celle saltate perché contenevano formule.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna: i comandi pubblici che la usano gestiscono -WhatIf e lavorano su una copia temporanea')]
    param(
        [Parameter(Mandatory)][string]$Percorso,
        [Parameter(Mandatory)][string]$ParteFoglio,
        [Parameter(Mandatory)][int]$IndiceFoglio,
        [Parameter(Mandatory)][object[]]$Modifiche,
        [int[]]$ColonneNumeriche = @(),
        [int]$RigaIntestazione = 1
    )
    $saltate = New-Object System.Collections.Generic.List[string]
    $zip = Open-HDZip -Percorso (Resolve-HDPercorso $Percorso) -Scrittura
    try {
        $doc = Read-HDXml -Zip $zip -Nome $ParteFoglio
        if (-not $doc) { throw "Foglio $ParteFoglio non trovato nel file." }
        $nsm = New-HDNsm $doc
        $sheetData = $doc.SelectSingleNode('/x:worksheet/x:sheetData', $nsm)
        $righe = @{}
        foreach ($r in $sheetData.SelectNodes('x:row', $nsm)) { $righe[[int]$r.GetAttribute('r')] = $r }
        $ultimaPrima = 0
        foreach ($n in $righe.Keys) {
            if ($n -gt $ultimaPrima -and $righe[$n].SelectSingleNode('x:c[x:v or x:is]', $nsm)) { $ultimaPrima = $n }
        }
        $stiliColonne = @{}
        foreach ($col in $doc.SelectNodes('/x:worksheet/x:cols/x:col', $nsm)) {
            if ($col.GetAttribute('style')) {
                for ($i = [int]$col.GetAttribute('min'); $i -le [int]$col.GetAttribute('max') -and $i -le 200; $i++) { $stiliColonne[$i] = $col.GetAttribute('style') }
            }
        }

        $primaNuova = [int]::MaxValue
        foreach ($m in ($Modifiche | Sort-Object { [int]$_.Riga })) {
            $n = [int]$m.Riga
            $row = $righe[$n]
            if (-not $row) {
                $row = New-HDElementoXml $doc 'row'
                $row.SetAttribute('r', [string]$n)
                $dopo = $null
                foreach ($r in $sheetData.SelectNodes('x:row', $nsm)) { if ([int]$r.GetAttribute('r') -gt $n) { $dopo = $r; break } }
                if ($dopo) { [void]$sheetData.InsertBefore($row, $dopo) } else { [void]$sheetData.AppendChild($row) }
                $righe[$n] = $row
            }
            if ($n -gt $ultimaPrima -and $n -lt $primaNuova) { $primaNuova = $n }
            $row.RemoveAttribute('spans')

            foreach ($chiave in @($m.Celle.Keys | Sort-Object { [int]$_ })) {
                $colonna = [int]$chiave
                $valore = ConvertTo-HDTestoXmlSicuro ([string]$m.Celle[$chiave])
                $riferimento = (ConvertTo-HDLettereColonna $colonna) + $n
                $cella = $null
                $dopo = $null
                foreach ($c in $row.SelectNodes('x:c', $nsm)) {
                    $rif = ConvertFrom-HDRiferimento $c.GetAttribute('r')
                    if (-not $rif) { continue }
                    if ($rif.Colonna -eq $colonna) { $cella = $c; break }
                    if ($rif.Colonna -gt $colonna) { $dopo = $c; break }
                }
                if ($cella -and $cella.SelectSingleNode('x:f', $nsm)) { $saltate.Add($riferimento); continue }
                if (-not $cella) {
                    $cella = New-HDElementoXml $doc 'c'
                    $cella.SetAttribute('r', $riferimento)
                    # stile: stessa colonna nella riga sopra più vicina, altrimenti quello della colonna
                    $stile = $null
                    for ($sopra = $n - 1; $sopra -gt $RigaIntestazione -and $sopra -ge $n - 200; $sopra--) {
                        if (-not $righe.ContainsKey($sopra)) { continue }
                        $vicina = $righe[$sopra].SelectSingleNode("x:c[@r='$((ConvertTo-HDLettereColonna $colonna) + $sopra)']", $nsm)
                        if ($vicina -and $vicina.GetAttribute('s')) { $stile = $vicina.GetAttribute('s'); break }
                    }
                    if (-not $stile -and $stiliColonne.ContainsKey($colonna)) { $stile = $stiliColonne[$colonna] }
                    if ($stile) { $cella.SetAttribute('s', $stile) }
                    if ($dopo) { [void]$row.InsertBefore($cella, $dopo) } else { [void]$row.AppendChild($cella) }
                }
                while ($cella.HasChildNodes) { [void]$cella.RemoveChild($cella.FirstChild) }
                foreach ($attributo in 't', 'cm', 'vm') { $cella.RemoveAttribute($attributo) }
                $numero = 0.0
                $eNumero = ($ColonneNumeriche -contains $colonna) -and [double]::TryParse($valore, [System.Globalization.NumberStyles]::Float, $script:HDInvariante, [ref]$numero)
                if ($eNumero) {
                    $v = New-HDElementoXml $doc 'v'
                    $v.InnerText = $numero.ToString('R', $script:HDInvariante)
                    [void]$cella.AppendChild($v)
                }
                else {
                    $cella.SetAttribute('t', 'inlineStr')
                    $is = New-HDElementoXml $doc 'is'
                    $t = New-HDElementoXml $doc 't'
                    $t.InnerText = $valore
                    if ($valore -ne $valore.Trim()) {
                        $spazi = $doc.CreateAttribute('xml', 'space', 'http://www.w3.org/XML/1998/namespace')
                        $spazi.Value = 'preserve'
                        [void]$t.Attributes.Append($spazi)
                    }
                    [void]$is.AppendChild($t)
                    [void]$cella.AppendChild($is)
                }
            }
        }

        # Dimensione, filtro automatico, tabelle
        $maxRiga = 0
        $maxRigaDati = 0
        $maxColonna = 1
        foreach ($c in $sheetData.SelectNodes('x:row/x:c', $nsm)) {
            $rif = ConvertFrom-HDRiferimento $c.GetAttribute('r')
            if ($rif) {
                if ($rif.Riga -gt $maxRiga) { $maxRiga = $rif.Riga }
                if ($rif.Colonna -gt $maxColonna) { $maxColonna = $rif.Colonna }
                if ($rif.Riga -gt $maxRigaDati -and $c.SelectSingleNode('x:v|x:is', $nsm)) { $maxRigaDati = $rif.Riga }
            }
        }
        $dimensione = $doc.SelectSingleNode('/x:worksheet/x:dimension', $nsm)
        if ($dimensione -and $maxRiga -gt 0) { $dimensione.SetAttribute('ref', ('A1:{0}{1}' -f (ConvertTo-HDLettereColonna $maxColonna), $maxRiga)) }

        if ($primaNuova -ne [int]::MaxValue) {
            $filtro = $doc.SelectSingleNode('/x:worksheet/x:autoFilter', $nsm)
            if ($filtro) {
                $nuovo = Expand-HDIntervallo -Intervallo $filtro.GetAttribute('ref') -PrimaNuova $primaNuova -UltimaRiga $maxRigaDati
                if ($nuovo) {
                    $filtro.SetAttribute('ref', $nuovo)
                    $info = Get-HDInfoWorkbook -Zip $zip
                    $wb = Read-HDXml -Zip $zip -Nome $info.ParteWorkbook
                    $nsmWb = New-HDNsm $wb
                    $nome = $wb.SelectSingleNode("/x:workbook/x:definedNames/x:definedName[@name='_xlnm._FilterDatabase' and @localSheetId='$IndiceFoglio']", $nsmWb)
                    if ($nome -and $nome.InnerText -match '\$(\d+)$') {
                        $nome.InnerText = $nome.InnerText -replace '\$\d+$', ('$' + $maxRigaDati)
                        Write-HDVoceZip -Zip $zip -Nome $info.ParteWorkbook -Testo $wb.OuterXml
                    }
                }
            }
            foreach ($rel in @(Get-HDRelazioni -Zip $zip -Parte $ParteFoglio)) {
                if ($rel.Tipo -notlike '*/table' -or $rel.Esterna) { continue }
                $tab = Read-HDXml -Zip $zip -Nome $rel.Parte
                if (-not $tab) { continue }
                $radice = $tab.DocumentElement
                if ([int]('0' + $radice.GetAttribute('totalsRowCount')) -gt 0) { continue }
                $nuovo = Expand-HDIntervallo -Intervallo $radice.GetAttribute('ref') -PrimaNuova $primaNuova -UltimaRiga $maxRigaDati
                if (-not $nuovo) { continue }
                $radice.SetAttribute('ref', $nuovo)
                $nsmTab = New-HDNsm $tab
                $filtroTab = $tab.SelectSingleNode('/x:table/x:autoFilter', $nsmTab)
                if ($filtroTab) { $filtroTab.SetAttribute('ref', $nuovo) }
                Write-HDVoceZip -Zip $zip -Nome $rel.Parte -Testo $tab.OuterXml
            }
        }

        Write-HDVoceZip -Zip $zip -Nome $ParteFoglio -Testo $doc.OuterXml
    }
    finally { $zip.Dispose() }
    , $saltate.ToArray()
}

function New-HDXlsx {
    <#
    Crea un .xlsx con un foglio: intestazione formattata e bloccata, filtro automatico,
    larghezze di colonna, colonne "da compilare a mano" evidenziate, elenchi a discesa e
    regole di evidenziazione.
    $Colonne: elenco di @{ Titolo; Larghezza; Manuale (bool); Testo (bool: formato Testo) }
    $Elenchi: @{ <indice colonna> = 'Valore1,Valore2' }
    $Evidenzia: @{ <indice colonna> = 'formula con {0} al posto della prima cella, es. AND(ISNUMBER({0}),{0}<15)' }
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna: New-Censimento gestisce -WhatIf')]
    param(
        [Parameter(Mandatory)][string]$Percorso,
        [Parameter(Mandatory)][string]$NomeFoglio,
        [Parameter(Mandatory)][object[]]$Colonne,
        [hashtable]$Elenchi = @{},
        [hashtable]$Evidenzia = @{},
        [int]$RigheRegole = 5000
    )
    Initialize-HDZip
    $esc = { param($s) [System.Security.SecurityElement]::Escape([string]$s) }
    $ultima = ConvertTo-HDLettereColonna $Colonne.Count
    $nomeFoglioXml = & $esc $NomeFoglio
    $nomeFoglioFormula = "'" + $NomeFoglio.Replace("'", "''") + "'"

    $cols = for ($i = 0; $i -lt $Colonne.Count; $i++) {
        $stile = 0
        if ($Colonne[$i].Testo) { $stile = 3 }
        '<col min="{0}" max="{0}" width="{1}" style="{2}" customWidth="1"/>' -f ($i + 1), ([double]$Colonne[$i].Larghezza).ToString($script:HDInvariante), $stile
    }
    $celle = for ($i = 0; $i -lt $Colonne.Count; $i++) {
        $stile = 1
        if ($Colonne[$i].Manuale) { $stile = 2 }
        '<c r="{0}1" s="{1}" t="inlineStr"><is><t>{2}</t></is></c>' -f (ConvertTo-HDLettereColonna ($i + 1)), $stile, (& $esc $Colonne[$i].Titolo)
    }
    $priorita = 0
    $condizionali = foreach ($k in ($Evidenzia.Keys | Sort-Object)) {
        $priorita++
        $lettera = ConvertTo-HDLettereColonna ([int]$k)
        '<conditionalFormatting sqref="{0}2:{0}{1}"><cfRule type="expression" dxfId="0" priority="{2}"><formula>{3}</formula></cfRule></conditionalFormatting>' -f $lettera, $RigheRegole, $priorita, (& $esc ($Evidenzia[$k] -f "$($lettera)2"))
    }
    $convalide = foreach ($k in ($Elenchi.Keys | Sort-Object)) {
        $lettera = ConvertTo-HDLettereColonna ([int]$k)
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="0" sqref="{0}2:{0}{1}"><formula1>{2}</formula1></dataValidation>' -f $lettera, $RigheRegole, (& $esc ('"' + $Elenchi[$k] + '"'))
    }
    $blocchiConvalide = ''
    if (@($convalide).Count -gt 0) { $blocchiConvalide = '<dataValidations count="{0}">{1}</dataValidations>' -f @($convalide).Count, (@($convalide) -join '') }

    $intestazioneXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $parti = [ordered]@{}
    $parti['[Content_Types].xml'] = $intestazioneXml + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    $parti['_rels/.rels'] = $intestazioneXml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $parti['xl/workbook.xml'] = $intestazioneXml + ('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="{0}" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">{1}!$A$1:${2}$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">{1}!$1:$1</definedName></definedNames></workbook>' -f $nomeFoglioXml, (& $esc $nomeFoglioFormula), $ultima)
    $parti['xl/_rels/workbook.xml.rels'] = $intestazioneXml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    $parti['xl/styles.xml'] = $intestazioneXml + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FF8EA9DB"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'
    $parti['xl/worksheets/sheet1.xml'] = $intestazioneXml + ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:{0}1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>{1}</cols><sheetData><row r="1" ht="32" customHeight="1">{2}</row></sheetData><autoFilter ref="A1:{0}1"/>{3}{4}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>' -f $ultima, (@($cols) -join ''), (@($celle) -join ''), (@($condizionali) -join ''), $blocchiConvalide)

    $p = Resolve-HDPercorso $Percorso
    [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $p))
    $flusso = [System.IO.File]::Open($p, 'CreateNew', 'ReadWrite', 'None')
    $zip = New-Object System.IO.Compression.ZipArchive($flusso, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    try { foreach ($nome in $parti.Keys) { Write-HDVoceZip -Zip $zip -Nome $nome -Testo $parti[$nome] } }
    finally { $zip.Dispose() }
}

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
    Legge i fogli: per ognuno Nome, Parte, Indice e Righe (numero riga -> hashtable colonna -> testo).
    -SoloFogli / -EscludiFogli: i fogli non richiesti tornano con Righe vuote (più veloce con fogli grandi).
    #>
    param([Parameter(Mandatory)][string]$Percorso, [string[]]$SoloFogli, [string[]]$EscludiFogli)
    $zip = Open-HDZip -Percorso (Resolve-HDPercorso $Percorso)
    $ns = $script:HDNsMain
    try {
        $info = Get-HDInfoWorkbook -Zip $zip
        $condivise = $info.Condivise
        foreach ($f in $info.Fogli) {
            $righe = New-Object 'System.Collections.Generic.SortedDictionary[int,hashtable]'
            $leggi = (-not $SoloFogli -or $SoloFogli -contains $f.Nome) -and -not ($EscludiFogli -contains $f.Nome)
            $doc = $null
            if ($leggi) { $doc = Read-HDXml -Zip $zip -Nome $f.Parte }
            $sheetData = $null
            if ($doc) {
                $nsm = New-HDNsm $doc
                $sheetData = $doc.SelectSingleNode('/x:worksheet/x:sheetData', $nsm)
            }
            if ($sheetData) {
                $numero = 0
                foreach ($row in $sheetData.ChildNodes) {
                    if ($row.LocalName -ne 'row') { continue }
                    $r = $row.GetAttribute('r')
                    if ($r) { $numero = [int]$r } else { $numero++ }
                    $celle = @{}
                    $colonna = 0
                    foreach ($c in $row.ChildNodes) {
                        if ($c.LocalName -ne 'c') { continue }
                        $rif = $c.GetAttribute('r')
                        if ($rif) {
                            $col = 0
                            foreach ($ch in $rif.ToCharArray()) {
                                $k = [int]$ch
                                if ($k -eq 36) { continue }
                                if ($k -ge 97) { $k -= 32 }
                                if ($k -lt 65 -or $k -gt 90) { break }
                                $col = $col * 26 + $k - 64
                            }
                            $colonna = $col
                        }
                        else { $colonna++ }
                        $valore = ''
                        $tipo = $c.GetAttribute('t')
                        if ($tipo -eq 's') {
                            $v = $c.Item('v', $ns)
                            $indice = 0
                            if ($v -and [int]::TryParse($v.InnerText, [ref]$indice) -and $indice -lt $condivise.Count) { $valore = $condivise[$indice] }
                        }
                        elseif ($tipo -eq 'inlineStr') {
                            $is = $c.Item('is', $ns)
                            if ($is) {
                                if ($is.InnerXml.Contains('rPh')) { $valore = Get-HDTestoNodo $is $nsm } else { $valore = $is.InnerText }
                            }
                        }
                        else {
                            $v = $c.Item('v', $ns)
                            if ($v) { $valore = $v.InnerText }
                        }
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
        $ultimaEsistente = 0
        foreach ($r in $sheetData.SelectNodes('x:row', $nsm)) {
            $nr = [int]$r.GetAttribute('r')
            $righe[$nr] = $r
            if ($nr -gt $ultimaEsistente) { $ultimaEsistente = $nr }
        }
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
                if ($n -lt $ultimaEsistente) {
                    foreach ($r in $sheetData.SelectNodes('x:row', $nsm)) { if ([int]$r.GetAttribute('r') -gt $n) { $dopo = $r; break } }
                }
                if ($dopo) { [void]$sheetData.InsertBefore($row, $dopo) } else { [void]$sheetData.AppendChild($row); $ultimaEsistente = $n }
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
                        if ($vicina) { $stile = $vicina.GetAttribute('s'); break }
                    }
                    if (-not $stile -and $n -eq $RigaIntestazione) {
                        # colonna nuova nell'intestazione: stesso stile della prima intestazione
                        $prima = $row.SelectSingleNode('x:c[1]', $nsm)
                        if ($prima) { $stile = $prima.GetAttribute('s') }
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
        foreach ($row in $sheetData.SelectNodes('x:row[x:c]', $nsm)) {
            $nr = [int]$row.GetAttribute('r')
            if ($nr -gt $maxRiga) { $maxRiga = $nr }
            $ultimaCella = $row.SelectSingleNode('x:c[last()]', $nsm)
            $rif = ConvertFrom-HDRiferimento $ultimaCella.GetAttribute('r')
            if ($rif -and $rif.Colonna -gt $maxColonna) { $maxColonna = $rif.Colonna }
            if ($nr -gt $maxRigaDati -and $row.SelectSingleNode('x:c[x:v or x:is]', $nsm)) { $maxRigaDati = $nr }
        }
        $dimensione = $doc.SelectSingleNode('/x:worksheet/x:dimension', $nsm)
        if ($dimensione -and $maxRiga -gt 0) { $dimensione.SetAttribute('ref', ('A1:{0}{1}' -f (ConvertTo-HDLettereColonna $maxColonna), $maxRiga)) }

        # Colonne aggiunte in fondo all'intestazione: il filtro automatico le comprende
        $filtroColonne = $doc.SelectSingleNode('/x:worksheet/x:autoFilter', $nsm)
        $rigaTitoli = $righe[$RigaIntestazione]
        if ($filtroColonne -and $rigaTitoli -and $filtroColonne.GetAttribute('ref') -match '^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$' -and [int]$Matches[2] -eq $RigaIntestazione) {
            $finePrima = (ConvertFrom-HDRiferimento ($Matches[3] + '1')).Colonna
            $ultimaTitolo = ConvertFrom-HDRiferimento $rigaTitoli.SelectSingleNode('x:c[last()]', $nsm).GetAttribute('r')
            if ($ultimaTitolo -and $ultimaTitolo.Colonna -gt $finePrima) {
                $lettera = ConvertTo-HDLettereColonna $ultimaTitolo.Colonna
                $filtroColonne.SetAttribute('ref', ('{0}{1}:{2}{3}' -f $Matches[1], $Matches[2], $lettera, $Matches[4]))
                $infoWb = Get-HDInfoWorkbook -Zip $zip
                $wbColonne = Read-HDXml -Zip $zip -Nome $infoWb.ParteWorkbook
                $nsmWbColonne = New-HDNsm $wbColonne
                $nomeFiltro = $wbColonne.SelectSingleNode("/x:workbook/x:definedNames/x:definedName[@name='_xlnm._FilterDatabase' and @localSheetId='$IndiceFoglio']", $nsmWbColonne)
                if ($nomeFiltro -and $nomeFiltro.InnerText -match ':\$[A-Z]+\$(\d+)$') {
                    $nomeFiltro.InnerText = $nomeFiltro.InnerText -replace ':\$[A-Z]+\$(\d+)$', (':$$' + $lettera + '$$$1')
                    Write-HDVoceZip -Zip $zip -Nome $infoWb.ParteWorkbook -Testo $wbColonne.OuterXml
                }
            }
        }

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

function Get-HDXmlFoglioNuovo {
    # XML di un foglio nuovo: intestazione (stili indicati), riga bloccata, filtro, larghezze,
    # formato Testo sulle colonne, elenchi a discesa ed evidenziazioni
    param(
        [Parameter(Mandatory)][object[]]$Colonne,
        [int]$StileIntestazione = 1, [int]$StileManuale = 2, [int]$StileTesto = 3, [int]$Dxf = 0,
        [hashtable]$Elenchi = @{}, [hashtable]$Evidenzia = @{}, [int]$RigheRegole = 5000, [switch]$Selezionato
    )
    $esc = { param($t) [System.Security.SecurityElement]::Escape([string]$t) }
    $ultima = ConvertTo-HDLettereColonna $Colonne.Count
    $cols = for ($i = 0; $i -lt $Colonne.Count; $i++) {
        $stile = 0
        if ($Colonne[$i].Testo) { $stile = $StileTesto }
        '<col min="{0}" max="{0}" width="{1}" style="{2}" customWidth="1"/>' -f ($i + 1), ([double]$Colonne[$i].Larghezza).ToString($script:HDInvariante), $stile
    }
    $celle = for ($i = 0; $i -lt $Colonne.Count; $i++) {
        $stile = $StileIntestazione
        if ($Colonne[$i].Manuale) { $stile = $StileManuale }
        '<c r="{0}1" s="{1}" t="inlineStr"><is><t>{2}</t></is></c>' -f (ConvertTo-HDLettereColonna ($i + 1)), $stile, (& $esc $Colonne[$i].Titolo)
    }
    $priorita = 0
    $condizionali = foreach ($k in ($Evidenzia.Keys | Sort-Object)) {
        $priorita++
        $lettera = ConvertTo-HDLettereColonna ([int]$k)
        '<conditionalFormatting sqref="{0}2:{0}{1}"><cfRule type="expression" dxfId="{4}" priority="{2}"><formula>{3}</formula></cfRule></conditionalFormatting>' -f $lettera, $RigheRegole, $priorita, (& $esc ($Evidenzia[$k] -f "$($lettera)2")), $Dxf
    }
    $convalide = foreach ($k in ($Elenchi.Keys | Sort-Object)) {
        $lettera = ConvertTo-HDLettereColonna ([int]$k)
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="0" sqref="{0}2:{0}{1}"><formula1>{2}</formula1></dataValidation>' -f $lettera, $RigheRegole, (& $esc ('"' + $Elenchi[$k] + '"'))
    }
    $blocchiConvalide = ''
    if (@($convalide).Count -gt 0) { $blocchiConvalide = '<dataValidations count="{0}">{1}</dataValidations>' -f @($convalide).Count, (@($convalide) -join '') }
    $scheda = ''
    if ($Selezionato) { $scheda = ' tabSelected="1"' }
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:{0}1"/><sheetViews><sheetView{5} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>{1}</cols><sheetData><row r="1" ht="32" customHeight="1">{2}</row></sheetData><autoFilter ref="A1:{0}1"/>{3}{4}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>' -f $ultima, (@($cols) -join ''), (@($celle) -join ''), (@($condizionali) -join ''), $blocchiConvalide, $scheda)
}

function Set-HDRigheFoglio {
    <#
    Riscrive tutto il contenuto di un foglio generato dal toolkit (intestazione + righe di testo).
    Mantiene colonne, stili delle colonne e impostazioni del foglio; aggiorna dimensione e filtro.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna su una copia temporanea')]
    param([Parameter(Mandatory)][string]$Percorso, [Parameter(Mandatory)][string]$ParteFoglio, [int]$IndiceFoglio, [Parameter(Mandatory)][string[]]$Intestazione, [object[]]$Righe = @(), [int[]]$ColonneNumeriche = @(), [double[]]$Larghezze = @())
    $zip = Open-HDZip -Percorso (Resolve-HDPercorso $Percorso) -Scrittura
    try {
        $doc = Read-HDXml -Zip $zip -Nome $ParteFoglio
        if (-not $doc) { throw "Foglio $ParteFoglio non trovato nel file." }
        $nsm = New-HDNsm $doc
        $sheetData = $doc.SelectSingleNode('/x:worksheet/x:sheetData', $nsm)
        $stileTitoli = ''
        $primaTitoli = $sheetData.SelectSingleNode('x:row[1]/x:c[1]', $nsm)
        if ($primaTitoli) { $stileTitoli = $primaTitoli.GetAttribute('s') }
        $altezza = $sheetData.SelectSingleNode('x:row[1]', $nsm)
        while ($sheetData.HasChildNodes) { [void]$sheetData.RemoveChild($sheetData.FirstChild) }
        $tutte = @(, [object[]]$Intestazione) + @($Righe)
        $numero = 0
        $maxColonna = $Intestazione.Count
        foreach ($valori in $tutte) {
            $numero++
            $row = New-HDElementoXml $doc 'row'
            $row.SetAttribute('r', [string]$numero)
            if ($numero -eq 1 -and $altezza -and $altezza.GetAttribute('ht')) {
                $row.SetAttribute('ht', $altezza.GetAttribute('ht'))
                $row.SetAttribute('customHeight', '1')
            }
            $valori = @($valori)
            if ($valori.Count -gt $maxColonna) { $maxColonna = $valori.Count }
            for ($i = 0; $i -lt $valori.Count; $i++) {
                $testo = (ConvertTo-HDTestoXmlSicuro ([string]$valori[$i])).Trim()
                if (-not $testo) { continue }
                $c = New-HDElementoXml $doc 'c'
                $c.SetAttribute('r', (ConvertTo-HDLettereColonna ($i + 1)) + $numero)
                if ($numero -eq 1 -and $stileTitoli) { $c.SetAttribute('s', $stileTitoli) }
                $valoreNumero = 0.0
                if ($numero -gt 1 -and $ColonneNumeriche -contains ($i + 1) -and [double]::TryParse($testo, [System.Globalization.NumberStyles]::Float, $script:HDInvariante, [ref]$valoreNumero)) {
                    $v = New-HDElementoXml $doc 'v'
                    $v.InnerText = $valoreNumero.ToString('R', $script:HDInvariante)
                    [void]$c.AppendChild($v)
                }
                else {
                    $c.SetAttribute('t', 'inlineStr')
                    $is = New-HDElementoXml $doc 'is'
                    $t = New-HDElementoXml $doc 't'
                    $t.InnerText = $testo
                    [void]$is.AppendChild($t)
                    [void]$c.AppendChild($is)
                }
                [void]$row.AppendChild($c)
            }
            [void]$sheetData.AppendChild($row)
        }
        $area = 'A1:{0}{1}' -f (ConvertTo-HDLettereColonna $maxColonna), $numero
        $dimensione = $doc.SelectSingleNode('/x:worksheet/x:dimension', $nsm)
        if ($dimensione) { $dimensione.SetAttribute('ref', $area) }
        $filtro = $doc.SelectSingleNode('/x:worksheet/x:autoFilter', $nsm)
        if ($filtro) {
            $filtro.SetAttribute('ref', $area)
            while ($filtro.HasChildNodes) { [void]$filtro.RemoveChild($filtro.FirstChild) }
            $info = Get-HDInfoWorkbook -Zip $zip
            $wb = Read-HDXml -Zip $zip -Nome $info.ParteWorkbook
            $nsmWb = New-HDNsm $wb
            $nome = $wb.SelectSingleNode("/x:workbook/x:definedNames/x:definedName[@name='_xlnm._FilterDatabase' and @localSheetId='$IndiceFoglio']", $nsmWb)
            if ($nome) {
                $foglioNome = $nome.InnerText.Substring(0, $nome.InnerText.LastIndexOf('!') + 1)
                $nome.InnerText = $foglioNome + ('$A$1:${0}${1}' -f (ConvertTo-HDLettereColonna $maxColonna), $numero)
                Write-HDVoceZip -Zip $zip -Nome $info.ParteWorkbook -Testo $wb.OuterXml
            }
        }
        # -Larghezze: larghezza di ogni colonna (serve quando cambia la disposizione del foglio)
        for ($i = 0; $i -lt $Larghezze.Count; $i++) {
            $n = [string]($i + 1)
            $col = $doc.SelectSingleNode("/x:worksheet/x:cols/x:col[@min='$n' and @max='$n']", $nsm)
            if ($col -and $Larghezze[$i] -gt 0) {
                $col.SetAttribute('width', $Larghezze[$i].ToString($script:HDInvariante))
                $col.SetAttribute('customWidth', '1')
            }
        }
        Write-HDVoceZip -Zip $zip -Nome $ParteFoglio -Testo $doc.OuterXml
    }
    finally { $zip.Dispose() }
}

function Get-HDNomeFoglioFormula {
    # Nome del foglio come si scrive in una formula: 'Censimento'
    param([string]$Nome)
    "'" + $Nome.Replace("'", "''") + "'"
}

function New-HDXlsx {
    <#
    Crea un .xlsx: il primo foglio con intestazione formattata e bloccata, filtro automatico,
    larghezze di colonna, colonne "da compilare a mano" evidenziate, elenchi a discesa e
    regole di evidenziazione; eventuali altri fogli (-AltriFogli) con intestazione e filtro.
    $Colonne: elenco di @{ Titolo; Larghezza; Manuale (bool); Testo (bool: formato Testo) }
    $Elenchi: @{ <indice colonna> = 'Valore1,Valore2' }
    $Evidenzia: @{ <indice colonna> = 'formula con {0} al posto della prima cella, es. AND(ISNUMBER({0}),{0}<15)' }
    $AltriFogli: elenco di @{ Nome; Colonne }
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna: New-Censimento gestisce -WhatIf')]
    param(
        [Parameter(Mandatory)][string]$Percorso,
        [Parameter(Mandatory)][string]$NomeFoglio,
        [Parameter(Mandatory)][object[]]$Colonne,
        [hashtable]$Elenchi = @{},
        [hashtable]$Evidenzia = @{},
        [int]$RigheRegole = 5000,
        [object[]]$AltriFogli = @()
    )
    Initialize-HDZip
    $esc = { param($t) [System.Security.SecurityElement]::Escape([string]$t) }
    $fogli = @([pscustomobject]@{ Nome = $NomeFoglio; Colonne = $Colonne; Evidenzia = $Evidenzia }) + @($AltriFogli | ForEach-Object {
            $ev = @{}
            if ($_.Evidenzia) { $ev = $_.Evidenzia }
            [pscustomobject]@{ Nome = [string]$_.Nome; Colonne = @($_.Colonne); Evidenzia = $ev; Nascosto = [bool]$_.Nascosto }
        })
    $overrides = ''
    $relazioni = ''
    $elencoFogli = ''
    $nomi = ''
    $xmlFogli = [ordered]@{}
    for ($i = 0; $i -lt $fogli.Count; $i++) {
        $n = $i + 1
        $f = $fogli[$i]
        $overrides += '<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $n
        $relazioni += '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $n
        $stato = ''
        if ($f.Nascosto) { $stato = ' state="hidden"' }
        $elencoFogli += '<sheet name="{0}" sheetId="{1}"{2} r:id="rId{1}"/>' -f (& $esc $f.Nome), $n, $stato
        $formula = & $esc (Get-HDNomeFoglioFormula $f.Nome)
        $nomi += '<definedName name="_xlnm._FilterDatabase" localSheetId="{0}" hidden="1">{1}!$A$1:${2}$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="{0}">{1}!$1:$1</definedName>' -f $i, $formula, (ConvertTo-HDLettereColonna $f.Colonne.Count)
        if ($i -eq 0) { $xmlFogli["xl/worksheets/sheet$n.xml"] = Get-HDXmlFoglioNuovo -Colonne $f.Colonne -Elenchi $Elenchi -Evidenzia $Evidenzia -RigheRegole $RigheRegole -Selezionato }
        else { $xmlFogli["xl/worksheets/sheet$n.xml"] = Get-HDXmlFoglioNuovo -Colonne $f.Colonne -Evidenzia $f.Evidenzia -RigheRegole $RigheRegole }
    }
    $rIdStili = $fogli.Count + 1
    $intestazioneXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $parti = [ordered]@{}
    $parti['[Content_Types].xml'] = $intestazioneXml + ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{0}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' -f $overrides)
    $parti['_rels/.rels'] = $intestazioneXml + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $parti['xl/workbook.xml'] = $intestazioneXml + ('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>{0}</sheets><definedNames>{1}</definedNames></workbook>' -f $elencoFogli, $nomi)
    $parti['xl/_rels/workbook.xml.rels'] = $intestazioneXml + ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{0}<Relationship Id="rId{1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' -f $relazioni, $rIdStili)
    $parti['xl/styles.xml'] = $intestazioneXml + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FF8EA9DB"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>'
    foreach ($k in $xmlFogli.Keys) { $parti[$k] = $xmlFogli[$k] }

    $p = Resolve-HDPercorso $Percorso
    [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $p))
    $flusso = [System.IO.File]::Open($p, 'CreateNew', 'ReadWrite', 'None')
    $zip = New-Object System.IO.Compression.ZipArchive($flusso, [System.IO.Compression.ZipArchiveMode]::Create, $false)
    try { foreach ($nome in $parti.Keys) { Write-HDVoceZip -Zip $zip -Nome $nome -Testo $parti[$nome] } }
    finally { $zip.Dispose() }
}

function Add-HDStiliFoglio {
    # Aggiunge agli stili di un file esistente quelli per l'intestazione e per le colonne di testo.
    # Restituisce gli indici (0 = stile normale se il file non ha stili da estendere).
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna su una copia temporanea')]
    param([Parameter(Mandatory)]$Zip, [Parameter(Mandatory)][string]$ParteWorkbook)
    $esito = [pscustomobject]@{ Intestazione = 0; Testo = 0 }
    $relStili = @(Get-HDRelazioni -Zip $Zip -Parte $ParteWorkbook | Where-Object { $_.Tipo -like '*/styles' }) | Select-Object -First 1
    if (-not $relStili) { return $esito }
    $st = Read-HDXml -Zip $Zip -Nome $relStili.Parte
    if (-not $st) { return $esito }
    $nsm = New-HDNsm $st
    $fonts = $st.SelectSingleNode('/x:styleSheet/x:fonts', $nsm)
    $fills = $st.SelectSingleNode('/x:styleSheet/x:fills', $nsm)
    $borders = $st.SelectSingleNode('/x:styleSheet/x:borders', $nsm)
    $xfs = $st.SelectSingleNode('/x:styleSheet/x:cellXfs', $nsm)
    if (-not ($fonts -and $fills -and $borders -and $xfs)) { return $esito }
    $aggiungi = {
        param($Contenitore, [string]$Xml)
        $frammento = $st.CreateDocumentFragment()
        $frammento.InnerXml = $Xml
        $nodo = $frammento.FirstChild
        [void]$Contenitore.AppendChild($nodo)
        $quanti = @($Contenitore.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }).Count
        $Contenitore.SetAttribute('count', [string]$quanti)
        $quanti - 1
    }
    $ns = $script:HDNsMain
    $font = & $aggiungi $fonts ('<font xmlns="{0}"><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' -f $ns)
    $fill = & $aggiungi $fills ('<fill xmlns="{0}"><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>' -f $ns)
    $bordo = & $aggiungi $borders ('<border xmlns="{0}"><left/><right/><top/><bottom style="thin"><color rgb="FF8EA9DB"/></bottom><diagonal/></border>' -f $ns)
    $esito.Intestazione = & $aggiungi $xfs ('<xf xmlns="{0}" numFmtId="0" fontId="{1}" fillId="{2}" borderId="{3}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' -f $ns, $font, $fill, $bordo)
    $esito.Testo = & $aggiungi $xfs ('<xf xmlns="{0}" numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' -f $ns)
    Write-HDVoceZip -Zip $Zip -Nome $relStili.Parte -Testo $st.OuterXml
    $esito
}

function Add-HDFoglioXlsx {
    <#
    Aggiunge in fondo a un file esistente un foglio nuovo con intestazione, riga bloccata e filtro.
    Non tocca gli altri fogli. Restituisce Parte e Indice del foglio creato.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzione interna: si usa su una copia temporanea')]
    param([Parameter(Mandatory)][string]$Percorso, [Parameter(Mandatory)][string]$NomeFoglio, [Parameter(Mandatory)][object[]]$Colonne, [switch]$Nascosto)
    $zip = Open-HDZip -Percorso (Resolve-HDPercorso $Percorso) -Scrittura
    try {
        $info = Get-HDInfoWorkbook -Zip $zip
        if (@($info.Fogli | Where-Object { $_.Nome -eq $NomeFoglio }).Count -gt 0) { throw "Il foglio '$NomeFoglio' esiste già." }
        $wb = Read-HDXml -Zip $zip -Nome $info.ParteWorkbook
        $nsm = New-HDNsm $wb
        $sheets = $wb.SelectSingleNode('/x:workbook/x:sheets', $nsm)
        $indice = $sheets.SelectNodes('x:sheet', $nsm).Count

        # Parte nuova accanto al workbook (di solito xl/worksheets/sheetN.xml)
        $cartellaWb = $info.ParteWorkbook -replace '[^/]+$', ''
        $n = 1
        while ($zip.GetEntry(('{0}worksheets/sheet{1}.xml' -f $cartellaWb, $n))) { $n++ }
        $parte = '{0}worksheets/sheet{1}.xml' -f $cartellaWb, $n

        # Relazione dal workbook
        $parteRels = Get-HDParteRelazione $info.ParteWorkbook
        $rels = Read-HDXml -Zip $zip -Nome $parteRels
        if (-not $rels) { throw 'Struttura del file non riconosciuta (relazioni del workbook mancanti).' }
        $nsmRels = New-HDNsm $rels
        $usati = @{}
        foreach ($r in $rels.SelectNodes('/p:Relationships/p:Relationship', $nsmRels)) { $usati[$r.GetAttribute('Id')] = $true }
        $k = 1
        while ($usati.ContainsKey("rIdHD$k")) { $k++ }
        $rid = "rIdHD$k"
        $rel = $rels.CreateElement('Relationship', $script:HDNsPkg)
        $rel.SetAttribute('Id', $rid)
        $rel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        $rel.SetAttribute('Target', ('worksheets/sheet{0}.xml' -f $n))
        [void]$rels.DocumentElement.AppendChild($rel)

        # Foglio nell'elenco del workbook
        $idMax = 0
        foreach ($sh in $sheets.SelectNodes('x:sheet', $nsm)) { if ([int]$sh.GetAttribute('sheetId') -gt $idMax) { $idMax = [int]$sh.GetAttribute('sheetId') } }
        $nuovo = New-HDElementoXml $wb 'sheet'
        $nuovo.SetAttribute('name', $NomeFoglio)
        $nuovo.SetAttribute('sheetId', [string]($idMax + 1))
        if ($Nascosto) { $nuovo.SetAttribute('state', 'hidden') }
        $prefisso = $wb.DocumentElement.GetPrefixOfNamespace($script:HDNsRel)
        if (-not $prefisso) {
            $prefisso = 'r'
            $wb.DocumentElement.SetAttribute('xmlns:r', $script:HDNsRel)
        }
        $attributo = $wb.CreateAttribute($prefisso, 'id', $script:HDNsRel)
        $attributo.Value = $rid
        [void]$nuovo.Attributes.Append($attributo)
        [void]$sheets.AppendChild($nuovo)

        # Nome per il filtro automatico (dopo sheets / functionGroups / externalReferences)
        $nomi = $wb.SelectSingleNode('/x:workbook/x:definedNames', $nsm)
        if (-not $nomi) {
            $nomi = New-HDElementoXml $wb 'definedNames'
            $dopo = $sheets
            foreach ($nome in 'functionGroups', 'externalReferences') {
                $trovato = $wb.SelectSingleNode("/x:workbook/x:$nome", $nsm)
                if ($trovato) { $dopo = $trovato }
            }
            [void]$wb.DocumentElement.InsertAfter($nomi, $dopo)
        }
        $definito = New-HDElementoXml $wb 'definedName'
        $definito.SetAttribute('name', '_xlnm._FilterDatabase')
        $definito.SetAttribute('localSheetId', [string]$indice)
        $definito.SetAttribute('hidden', '1')
        $definito.InnerText = '{0}!$A$1:${1}$1' -f (Get-HDNomeFoglioFormula $NomeFoglio), (ConvertTo-HDLettereColonna $Colonne.Count)
        [void]$nomi.AppendChild($definito)

        # Tipo della parte
        $tipi = Read-HDXml -Zip $zip -Nome '[Content_Types].xml'
        $override = $tipi.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
        $override.SetAttribute('PartName', '/' + $parte)
        $override.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        [void]$tipi.DocumentElement.AppendChild($override)

        $stili = Add-HDStiliFoglio -Zip $zip -ParteWorkbook $info.ParteWorkbook
        Write-HDVoceZip -Zip $zip -Nome $parte -Testo (Get-HDXmlFoglioNuovo -Colonne $Colonne -StileIntestazione $stili.Intestazione -StileManuale $stili.Intestazione -StileTesto $stili.Testo)
        Write-HDVoceZip -Zip $zip -Nome $parteRels -Testo $rels.OuterXml
        Write-HDVoceZip -Zip $zip -Nome $info.ParteWorkbook -Testo $wb.OuterXml
        Write-HDVoceZip -Zip $zip -Nome '[Content_Types].xml' -Testo $tipi.OuterXml
    }
    finally { $zip.Dispose() }
    [pscustomobject]@{ Parte = $parte; Indice = $indice }
}

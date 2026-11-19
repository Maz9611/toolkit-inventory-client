# Censimento: prodotti Autodesk (AutoCAD, Revit, Inventor...) con versione, tipo di licenza, account
# Autodesk, ultimo avvio e seriale. Foglio "Autodesk": una riga per prodotto installato su ogni PC.
# Funzioni pure: i dati grezzi li raccoglie $script:HDRaccoltaCensimento (blocco Autodesk).

$script:HDFoglioAutodesk = 'Autodesk'

$script:HDColonneAutodesk = @(
    New-HDColonnaFoglio 'Hostname' 'Hostname' 18
    New-HDColonnaFoglio 'Utilizzatore' 'Utilizzatore' 20
    New-HDColonnaFoglio 'Prodotto' 'Prodotto' 28
    New-HDColonnaFoglio 'Versione' 'Versione' 16
    New-HDColonnaFoglio 'Licenza' 'Licenza' 40
    New-HDColonnaFoglio 'Account' 'Account Autodesk' 36
    New-HDColonnaFoglio 'UltimoAvvio' 'Ultimo avvio' 16
    New-HDColonnaFoglio 'Installato' 'Installato il' 13
    New-HDColonnaFoglio 'Seriale' 'Seriale' 28
    New-HDColonnaFoglio 'Note' 'Note' 30
    New-HDColonnaFoglio 'Rilevato' 'Rilevato il' 17
)

# Prodotti principali (le voci "Autodesk ..." che sono componenti, librerie o servizi non contano)
$script:HDProdottiAutodesk = '(?i)^(Autodesk\s+)?(AutoCAD|Revit|Inventor|3ds Max|Maya|Civil 3D|Navisworks|Advance Steel|Vault|Fusion|Robot Structural Analysis|InfraWorks|ReCap|Vehicle Tracking|Alias|Factory Design|Moldflow|Point Layout|Structural Bridge Design|FormIt|Plant 3D|Map 3D|Raster Design)\b'
$script:HDComponentiAutodesk = '(?i)Language Pack|Language|Update|Hotfix|Service Pack|Material Library|Object Enabler|Add-?in|Plug-?in|Content|Template|Help|Single Sign|Licensing|Genuine|Access\b|Desktop App|Featured Apps|Save to Web|App Manager|Connector|Shared|Component|Runtime|Viewer|TrueView|Batch|Identity|Personali[sz]ation|Libraries|Library|Interop|Migration|Network License|Performance Reporting|Export|Import|Cloud|Docs|Collaboration|Interoperability|Exchange|Tool|Catalog|Browser|Utilit|Server|Manager'

# Prefissi delle chiavi di prodotto (le prime 3 cifre: "001P1" = AutoCAD 2024)
$script:HDChiaviAutodesk = [ordered]@{
    '001' = 'AutoCAD'; '057' = 'AutoCAD LT'; '185' = 'AutoCAD Architecture'; '206' = 'AutoCAD Mechanical'
    '225' = 'AutoCAD Electrical'; '235' = 'AutoCAD MEP'; '129' = 'AutoCAD Map 3D'; '426' = 'AutoCAD Plant 3D'
    '237' = 'Civil 3D'; '829' = 'Revit'; '828' = 'Revit LT'; '797' = 'Inventor Professional'; '128' = '3ds Max'
    '657' = 'Maya'; '507' = 'Navisworks Manage'; '959' = 'Advance Steel'
}

# Programma che avvia ogni famiglia di prodotti (per l'ultimo avvio dai file di Prefetch)
$script:HDEseguibiliAutodesk = [ordered]@{
    'AutoCAD LT' = 'ACADLT'; 'AutoCAD' = 'ACAD'; 'Civil 3D' = 'ACAD'; 'Advance Steel' = 'ACAD'; 'Revit' = 'REVIT'
    'Inventor' = 'INVENTOR'; '3ds Max' = '3DSMAX'; 'Maya' = 'MAYA'; 'Navisworks' = 'ROAMER'
}

# Seriali generici usati dalle copie non originali (mai dati da Autodesk a un cliente)
$script:HDSerialiGenericiAutodesk = '^(666-69696969|667-98989898|400-45454545)$'

function ConvertFrom-HDLicenzeAutodesk {
    # Uscita di "AdskLicensingInstHelper list" (JSON) -> prodotti registrati con la licenza scelta:
    # Chiave e Anno del prodotto, ChiaveScelta (la licenza usata: può essere quella di una Collection),
    # Modalita (USER, STANDALONE, NETWORK o vuota) e Server di licenza. Formato non documentato: il JSON
    # può avere testo prima e dopo, i prodotti possono stare dentro un oggetto, i nomi dei campi cambiare.
    param([string]$Testo)
    $t = ([string]$Testo).Replace([string][char]0, '').Trim([char]0xFEFF)
    $voci = $null
    $inizi = @(@(0) + @([regex]::Matches($t, '(?m)^\s*[\[{]') | ForEach-Object { $_.Index }) | Select-Object -Unique)
    foreach ($i in $inizi) {
        $pezzo = $t.Substring($i).Trim()
        $fine = [math]::Max($pezzo.LastIndexOf(']'), $pezzo.LastIndexOf('}'))
        if ($fine -lt 0) { continue }
        try { $voci = $pezzo.Substring(0, $fine + 1) | ConvertFrom-Json -ErrorAction Stop; break }
        catch { $voci = $null }
    }
    if ($null -eq $voci) { return }
    # i prodotti sono gli oggetti con una chiave di prodotto, ovunque siano nel JSON
    $prodotti = New-Object System.Collections.Generic.List[object]
    $cerca = {
        param($x, [int]$Livello)
        if ($null -eq $x -or $Livello -gt 6) { return }
        if ($x -is [System.Array]) { foreach ($e in $x) { & $cerca $e ($Livello + 1) }; return }
        if ($x -is [System.Management.Automation.PSCustomObject]) {
            $nomi = @($x.PSObject.Properties | ForEach-Object { $_.Name })
            if (@($nomi | Where-Object { $_ -match '^(def_|sel_)?prod(uct)?_?key$' }).Count -gt 0) { $prodotti.Add($x); return }
            foreach ($p in $x.PSObject.Properties) { & $cerca $p.Value ($Livello + 1) }
        }
    }
    & $cerca $voci 0
    $campo = {
        param($o, [string[]]$Nomi)
        foreach ($n in $Nomi) { if ($o.PSObject.Properties[$n] -and "$($o.$n)" -ne '') { return $o.$n } }
        $null
    }
    foreach ($v in $prodotti) {
        $chiave = [string](& $campo $v @('def_prod_key', 'prod_key', 'product_key', 'productKey', 'sel_prod_key'))
        $versione = [string](& $campo $v @('def_prod_ver', 'prod_ver', 'product_version', 'productVersion', 'sel_prod_ver'))
        $anno = ''
        if ($versione -match '^(\d{4})') { $anno = $Matches[1] }
        $serverGrezzi = @(& $campo $v @('sel_srv_names', 'lic_servers', 'server_names', 'servers'))
        $server = @($serverGrezzi | Where-Object { $_ } | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.name) { [string]$_.name } } | Select-Object -Unique)
        [pscustomobject]@{
            Chiave       = $chiave
            Anno         = $anno
            ChiaveScelta = [string](& $campo $v @('sel_prod_key', 'selected_prod_key'))
            Modalita     = ([string](& $campo $v @('sel_mode', 'lic_mode', 'license_mode', 'licenseMode', 'mode'))).ToUpperInvariant()
            Server       = ($server -join ', ')
        }
    }
}

function Get-HDNotaAutodesk {
    # Note sulla lettura dei dati Autodesk (per il log della prova): quelle della raccolta e, se la risposta
    # di AdskLicensingInstHelper non si è capita, il suo inizio (per correggere la lettura)
    param($Autodesk)
    if ($null -eq $Autodesk) { return '' }
    $note = @($Autodesk.Note | Where-Object { $_ })
    $testo = [string]$Autodesk.Licenze
    if ($testo.Trim() -and @(ConvertFrom-HDLicenzeAutodesk $testo).Count -eq 0) {
        $breve = ($testo -replace '\s+', ' ').Trim()
        $note += 'risposta di AdskLicensingInstHelper non capita, inizia con: ' + $breve.Substring(0, [math]::Min(300, $breve.Length))
    }
    elseif (-not $testo.Trim() -and @($note | Where-Object { $_ -like '*non trovato*' }).Count -eq 0) {
        $note += 'AdskLicensingInstHelper non ha scritto niente'
    }
    $note -join '; '
}

function ConvertFrom-HDAccountAutodesk {
    # LoginState.xml di ogni profilo di Windows -> account Autodesk (email) con cui si è entrati
    param([object[]]$File)
    foreach ($f in @($File | Where-Object { $_ -and $_.Testo })) {
        if ([string]$f.Testo -match '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}') {
            [pscustomobject]@{ Utente = [string]$f.Utente; Email = $Matches[0].ToLowerInvariant() }
        }
    }
}

function Select-HDProdottiAutodesk {
    # Elenco dei programmi installati -> prodotti Autodesk principali: Nome ("AutoCAD 2024"), Famiglia
    # ("AutoCAD"), Anno, Versione e data di installazione. Niente componenti, librerie e servizi.
    param($Programmi)
    $visti = @{}
    foreach ($p in @($Programmi | Where-Object { $_ -and $_.DisplayName })) {
        $nome = ([string]$p.DisplayName).Trim()
        $editore = [string]$p.Publisher
        if ($nome -notmatch $script:HDProdottiAutodesk) { continue }
        if ($editore -and $editore -notmatch '(?i)autodesk') { continue }
        if ($nome -match $script:HDComponentiAutodesk) { continue }
        if ([string]$p.SystemComponent -eq '1' -or $p.ParentKeyName) { continue }
        $pulito = ($nome -replace '^(?i)Autodesk\s+', '' -replace '\s+-\s+[^-]+$', '').Trim()
        $anno = ''
        if ($pulito -match '\b((?:19|20)\d\d)\b') { $anno = $Matches[1] }
        $famiglia = ($pulito -replace '\s*\b(?:19|20)\d\d\b.*$', '').Trim()
        $chiave = ConvertTo-HDChiave $pulito
        if ($visti.ContainsKey($chiave)) { continue }
        $visti[$chiave] = $true
        [pscustomobject]@{
            Nome       = $pulito
            Famiglia   = $famiglia
            Anno       = $anno
            Versione   = ([string]$p.DisplayVersion).Trim()
            Installato = (ConvertFrom-HDDataInstallazione ([string]$p.InstallDate))
        }
    }
}

function Format-HDLicenzaAutodesk {
    # Licenza di un prodotto in parole: dalla licenza registrata (2020 e successivi) o, per le versioni
    # vecchie, da seriale e server FlexNet
    param($Licenza, [string]$Anno, [string]$Rete)
    if ($Licenza) {
        $testo = switch ($Licenza.Modalita) {
            'USER' { 'Utente con nome (abbonamento, account Autodesk)'; break }
            'STANDALONE' { 'Singola con seriale'; break }
            'NETWORK' { 'Di rete'; break }
            '' { 'Nessuna licenza scelta'; break }
            default { 'Licenza di tipo ' + $Licenza.Modalita }
        }
        if ($Licenza.Modalita -eq 'NETWORK') {
            $server = $Licenza.Server
            if (-not $server) { $server = $Rete }
            if ($server) { $testo += ' (server {0})' -f $server }
        }
        $scelta = [string]$Licenza.ChiaveScelta
        if ($scelta -and $Licenza.Chiave -and $scelta.Length -ge 3 -and $scelta.Substring(0, 3) -ne $Licenza.Chiave.Substring(0, [math]::Min(3, $Licenza.Chiave.Length))) {
            $altro = $script:HDChiaviAutodesk[$scelta.Substring(0, 3)]
            if ($altro) { $testo += ', con la licenza di ' + $altro }
            else { $testo += ', con la licenza di un''altra offerta (chiave {0}, es. una Collection)' -f $scelta }
        }
        return $testo
    }
    if ($Anno -and [int]$Anno -le 2019) {
        if ($Rete) { return 'Di rete FlexNet (server {0})' -f $Rete }
        return 'Singola con seriale (versione 2019 o precedente)'
    }
    ''
}

function ConvertTo-HDAutodeskPC {
    # Dati grezzi -> righe del foglio Autodesk per questo PC: Prodotto, Versione, Licenza, Account,
    # UltimoAvvio, Installato, Seriale. $null se i dati Autodesk non sono stati raccolti.
    param($Programmi, $Autodesk, [datetime]$Adesso = (Get-Date))
    if ($null -eq $Autodesk) { return $null }
    $prodotti = @(Select-HDProdottiAutodesk $Programmi)
    if ($prodotti.Count -eq 0) { return , @() }
    $licenze = @(ConvertFrom-HDLicenzeAutodesk $Autodesk.Licenze)
    $usate = @{}
    $account = @(ConvertFrom-HDAccountAutodesk @($Autodesk.Account))
    $testoAccount = (@($account | ForEach-Object { '{0} ({1})' -f $_.Email, $_.Utente } | Select-Object -Unique) -join '; ')
    $rete = ((@($Autodesk.Rete | Where-Object { $_ }) | ForEach-Object { ([string]$_).Trim().TrimStart('@') } | Select-Object -Unique) -join ', ')
    $avvii = @{}
    if ($null -ne $Autodesk.Avvii) {
        foreach ($a in @($Autodesk.Avvii | Where-Object { $_ -and $_.Data })) {
            $d = [datetime]$a.Data
            $k = ([string]$a.Programma).ToUpperInvariant()
            if (-not $avvii.ContainsKey($k) -or $d -gt $avvii[$k]) { $avvii[$k] = $d }
        }
    }
    # programma che avvia ogni prodotto (il primo che corrisponde: "AutoCAD LT" prima di "AutoCAD")
    $exePer = @{}
    foreach ($p in $prodotti) {
        $exePer[$p.Nome] = ''
        foreach ($f in $script:HDEseguibiliAutodesk.Keys) { if ($p.Famiglia -like "$f*") { $exePer[$p.Nome] = $script:HDEseguibiliAutodesk[$f]; break } }
    }
    $righe = foreach ($p in $prodotti) {
        # licenza registrata: stessa famiglia (dal prefisso della chiave) e stesso anno
        $licenza = $null
        $prefisso = @($script:HDChiaviAutodesk.Keys | Where-Object { $script:HDChiaviAutodesk[$_] -eq $p.Famiglia }) | Select-Object -First 1
        if ($prefisso) { $licenza = @($licenze | Where-Object { $_.Anno -eq $p.Anno -and $_.Chiave -like "$prefisso*" }) | Select-Object -First 1 }
        if (-not $licenza -and $p.Anno) {
            $stessoAnno = @($licenze | Where-Object { $_.Anno -eq $p.Anno -and -not $usate.ContainsKey($_.Chiave) })
            if ($stessoAnno.Count -eq 1) { $licenza = $stessoAnno[0] }
        }
        if ($licenza) { $usate[$licenza.Chiave] = $true }
        $testoLicenza = Format-HDLicenzaAutodesk -Licenza $licenza -Anno $p.Anno -Rete $rete
        if (-not $testoLicenza) { $testoLicenza = 'Non letta' }

        # seriale (versioni con licenza a seriale): dalle chiavi di registro di AutoCAD con lo stesso nome
        $seriale = ''
        foreach ($s in @($Autodesk.Seriali | Where-Object { $_ -and $_.Seriale })) {
            $nomeRegistro = ([string]$s.Prodotto -replace '^(?i)Autodesk\s+', '').Trim()
            if ((ConvertTo-HDChiave $nomeRegistro) -eq (ConvertTo-HDChiave $p.Nome)) { $seriale = ([string]$s.Seriale).Trim(); break }
        }
        if ($seriale -match $script:HDSerialiGenericiAutodesk) { $seriale += ' (seriale delle copie non originali)' }

        # ultimo avvio dal Prefetch (lo stesso programma avvia tutte le versioni installate della famiglia)
        $ultimo = ''
        if ($null -eq $Autodesk.Avvii) { $ultimo = 'Serve avviare come amministratore' }
        else {
            $exe = $exePer[$p.Nome]
            if ($exe -and $avvii.ContainsKey($exe)) {
                $ultimo = $avvii[$exe].ToString('dd/MM/yyyy', $script:HDInvariante)
                $stessi = @($prodotti | Where-Object { $exePer[$_.Nome] -eq $exe }).Count
                if ($stessi -gt 1) { $ultimo += (' (uno dei {0} prodotti con lo stesso programma)' -f $stessi) }
            }
            elseif ($exe) { $ultimo = 'Mai (dall''ultima pulizia di Windows)' }
        }
        [pscustomobject]@{
            Prodotto    = $p.Nome
            Versione    = $p.Versione
            Licenza     = $testoLicenza
            Account     = $testoAccount
            UltimoAvvio = $ultimo
            Installato  = $p.Installato
            Seriale     = $seriale
        }
    }
    , @($righe)
}

function Get-HDRigheAutodesk {
    # Foglio Autodesk riscritto: le righe degli altri PC restano, quelle dei PC appena censiti si
    # sostituiscono. Le Note scritte a mano restano (stesso PC e stesso prodotto).
    # $Nuovi: un elemento per PC censito con Hostname, SerialNumber, Utilizzatore, Rilevato e Righe.
    param([object[]]$Esistenti, [object[]]$Nuovi)
    $idPc = {
        param($Seriale, $Nome)
        $s = Get-HDNumeroSerie $Seriale
        if ($s) { return 'S:' + $s.ToUpperInvariant() }
        'H:' + (ConvertTo-HDNomeHost $Nome)
    }
    $note = @{}
    $tenute = New-Object System.Collections.Generic.List[object]
    $censiti = @{}
    foreach ($n in @($Nuovi | Where-Object { $_ })) {
        $censiti[(& $idPc $n.SerialNumber $n.Hostname)] = $true
        $censiti['H:' + (ConvertTo-HDNomeHost $n.Hostname)] = $true
    }
    foreach ($r in @($Esistenti | Where-Object { $_ })) {
        $v = $r.Valori
        $id = & $idPc $v['SerialNumber'] $v['Hostname']
        if ($v['Note']) { $note[$id + '|' + (ConvertTo-HDChiave $v['Prodotto'])] = [string]$v['Note'] }
        if ($censiti.ContainsKey($id) -or $censiti.ContainsKey('H:' + (ConvertTo-HDNomeHost $v['Hostname']))) { continue }
        $tenute.Add($v)
    }
    foreach ($n in @($Nuovi | Where-Object { $_ })) {
        $id = & $idPc $n.SerialNumber $n.Hostname
        foreach ($p in @($n.Righe | Where-Object { $_ })) {
            $nota = $note[$id + '|' + (ConvertTo-HDChiave $p.Prodotto)]
            if (-not $nota) { $nota = $note['H:' + (ConvertTo-HDNomeHost $n.Hostname) + '|' + (ConvertTo-HDChiave $p.Prodotto)] }
            $tenute.Add(@{
                    Hostname = [string]$n.Hostname; SerialNumber = [string]$n.SerialNumber; Utilizzatore = [string]$n.Utilizzatore
                    Prodotto = $p.Prodotto; Versione = $p.Versione; Licenza = $p.Licenza; Account = $p.Account
                    UltimoAvvio = $p.UltimoAvvio; Installato = $p.Installato; Seriale = $p.Seriale; Note = [string]$nota; Rilevato = [string]$n.Rilevato
                })
        }
    }
    $tenute | Sort-Object { [string]$_['Prodotto'] }, { [string]$_['Hostname'] }
}

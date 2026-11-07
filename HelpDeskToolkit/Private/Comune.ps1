# Funzioni comuni del toolkit (private: non esportate dal modulo).
# Regola d'oro: compatibili con Windows PowerShell 5.1 (niente ??, ?:, &&, Join-String, ecc.).

$script:HDCultura = [System.Globalization.CultureInfo]::GetCultureInfo('it-IT')
$script:HDInvariante = [System.Globalization.CultureInfo]::InvariantCulture

#region Configurazione e cartelle

function Get-HDConfigPredefinita {
    @{
        CartellaReport  = ''
        ProponiAmministratore = $true
        Censimento      = @{
            Cartella       = ''
            ModelloExcel   = ''
            RilevaCynet    = 'Cynet'
            AccountTecnici = @('admin', 'administrator', 'amministratore', 'tecnico')
            BackupDaTenere = 30
            DeviceTypeClient = 'PC CLIENT'
            DeviceTypeServer = 'SERVER'
            DominioAssente   = 'No'
        }
        Soglie          = @{
            DiscoAttenzionePct            = 10
            DiscoAttenzioneGB             = 20
            DiscoCriticoPct               = 5
            DiscoCriticoGB                = 10
            UptimeGiorni                  = 14
            RamUsoPct                     = 90
            AggiornamentiGiorniAttenzione = 45
            AggiornamentiGiorniCritico    = 90
            FirmeAntivirusGiorni          = 7
            ScartoOrarioMinuti            = 5
        }
    }
}

function Merge-HDHashtable {
    param([hashtable]$Base, [hashtable]$Sovrascrivi)
    $risultato = @{}
    foreach ($k in $Base.Keys) { $risultato[$k] = $Base[$k] }
    if ($Sovrascrivi) {
        foreach ($k in $Sovrascrivi.Keys) {
            if ($risultato[$k] -is [hashtable] -and $Sovrascrivi[$k] -is [hashtable]) {
                $risultato[$k] = Merge-HDHashtable -Base $risultato[$k] -Sovrascrivi $Sovrascrivi[$k]
            }
            else {
                $risultato[$k] = $Sovrascrivi[$k]
            }
        }
    }
    $risultato
}

function Get-HDConfig {
    # Legge config.psd1 (nella cartella del toolkit, oppure il file indicato in
    # $env:HELPDESK_TOOLKIT_CONFIG) e lo unisce ai valori predefiniti.
    [CmdletBinding()]
    param([string]$Percorso)
    if (-not $Percorso) {
        if ($env:HELPDESK_TOOLKIT_CONFIG) { $Percorso = $env:HELPDESK_TOOLKIT_CONFIG }
        else { $Percorso = Join-Path $script:HDToolkitRadice 'config.psd1' }
    }
    $cfg = Get-HDConfigPredefinita
    if (Test-Path -LiteralPath $Percorso) {
        try {
            $utente = Import-PowerShellDataFile -LiteralPath $Percorso -ErrorAction Stop
            $cfg = Merge-HDHashtable -Base $cfg -Sovrascrivi $utente
        }
        catch {
            Write-Warning "config.psd1 non valido ($Percorso): $($_.Exception.Message). Uso i valori predefiniti."
        }
    }
    $cfg
}

function Resolve-HDPercorso {
    # Percorso assoluto (i metodi .NET non conoscono la cartella corrente di PowerShell)
    param([Parameter(Mandatory)][string]$Percorso)
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Percorso)
}

function Test-HDCartellaScrivibile {
    param([Parameter(Mandatory)][string]$Percorso)
    try {
        [void][System.IO.Directory]::CreateDirectory($Percorso)
        $prova = Join-Path $Percorso ('.prova-scrittura-' + [guid]::NewGuid().ToString('N') + '.tmp')
        [System.IO.File]::WriteAllText($prova, 'ok')
        Remove-Item -LiteralPath $prova -Force -ErrorAction SilentlyContinue
        $true
    }
    catch { $false }
}

function Get-HDCartella {
    # Cartella dove salvare report o note. Se non è scrivibile (es. chiavetta protetta,
    # cartella in sola lettura) ripiega su %TEMP%\HelpDeskToolkit.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Report', 'Censimento')][string]$Tipo,
        [string]$Specificata
    )
    $cfg = Get-HDConfig
    $configurata = $cfg['CartellaReport']
    if ($Tipo -eq 'Censimento') { $configurata = $cfg.Censimento.Cartella }
    if ($Specificata) { $percorso = $Specificata }
    elseif ($configurata) { $percorso = [Environment]::ExpandEnvironmentVariables([string]$configurata) }
    else { $percorso = Join-Path $script:HDToolkitRadice $Tipo }
    $percorso = Resolve-HDPercorso $percorso
    if (Test-HDCartellaScrivibile $percorso) { return $percorso }
    if ($Specificata) { throw "Impossibile scrivere nella cartella '$percorso'." }
    $alternativa = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'HelpDeskToolkit') $Tipo
    Write-Warning "Non posso scrivere in '$percorso' (sola lettura?): uso '$alternativa'."
    if (-not (Test-HDCartellaScrivibile $alternativa)) { throw "Impossibile scrivere anche in '$alternativa'." }
    $alternativa
}

function Get-HDPercorsoUnico {
    # Evita di sovrascrivere: Rapportino.html -> Rapportino-2.html -> ...
    param([string]$Cartella, [string]$NomeBase, [string]$Estensione)
    $p = Join-Path $Cartella ($NomeBase + $Estensione)
    $n = 2
    while (Test-Path -LiteralPath $p) {
        $p = Join-Path $Cartella ('{0}-{1}{2}' -f $NomeBase, $n, $Estensione)
        $n++
    }
    $p
}

#endregion

#region Piattaforma

function Test-HDWindows {
    if ($PSVersionTable.PSVersion.Major -le 5) { return $true }
    return [bool](Get-Variable -Name IsWindows -ValueOnly -ErrorAction SilentlyContinue)
}

function Test-HDAmministratore {
    if (-not (Test-HDWindows)) { return $false }
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { $false }
}

function Get-HDNomePC {
    if ($env:COMPUTERNAME) { return $env:COMPUTERNAME }
    [Environment]::MachineName
}

function Get-HDCim {
    # Unico punto di accesso a CIM/WMI: nei test viene sostituito con un mock.
    param(
        [Parameter(Mandatory)][string]$Classe,
        [string]$Namespace = 'root/cimv2',
        [string]$Filtro,
        [string]$ComputerName
    )
    $p = @{ ClassName = $Classe; Namespace = $Namespace; ErrorAction = 'Stop' }
    if ($Filtro) { $p.Filter = $Filtro }
    if ($ComputerName) { $p.ComputerName = $ComputerName }
    Get-CimInstance @p
}

#endregion

#region Testo, numeri e date

function ConvertTo-HDSenzaAccenti {
    param([string]$Testo)
    if (-not $Testo) { return '' }
    $norm = $Testo.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

function ConvertTo-HDChiave {
    # "  Modalità  Intervento " -> "modalita intervento"
    param([string]$Testo)
    $s = (ConvertTo-HDSenzaAccenti $Testo).ToLowerInvariant()
    $s = $s -replace "[^a-z0-9\- ]", ' '
    ($s -replace '\s+', ' ').Trim()
}

function ConvertTo-HDNomeFile {
    param([string]$Testo, [int]$Massimo = 40)
    $s = (ConvertTo-HDSenzaAccenti $Testo) -replace '[^A-Za-z0-9]+', '-'
    $s = $s.Trim('-')
    if ($s.Length -gt $Massimo) { $s = $s.Substring(0, $Massimo).Trim('-') }
    if (-not $s) { return 'senza-nome' }
    $s
}

function Format-HDNumero {
    param([double]$Valore, [int]$Decimali = 1)
    $Valore.ToString("N$Decimali", $script:HDCultura)
}

function Format-HDQuantita {
    # 1 -> "1 stampante", 3 -> "3 stampanti"
    param([int]$Numero, [string]$Singolare, [string]$Plurale)
    if ($Numero -eq 1) { return "1 $Singolare" }
    "$Numero $Plurale"
}

function Join-HDProduttoreModello {
    # "HP" + "HP ProDesk 400 G7" -> "HP ProDesk 400 G7" (niente "HP HP")
    param([string]$Produttore, [string]$Modello)
    $p = ([string]$Produttore).Trim()
    $m = ([string]$Modello).Trim()
    if (-not $p) { return $m }
    $primaParola = ($p -split '\s+')[0]
    if ($m.StartsWith($primaParola, [System.StringComparison]::OrdinalIgnoreCase)) { return $m }
    ("$p $m").Trim()
}

function Format-HDByte {
    param([double]$Byte)
    if ($Byte -ge 1TB) { return '{0} TB' -f (Format-HDNumero ($Byte / 1TB) 2) }
    if ($Byte -ge 1GB) { return '{0} GB' -f (Format-HDNumero ($Byte / 1GB) 1) }
    if ($Byte -ge 1MB) { return '{0} MB' -f (Format-HDNumero ($Byte / 1MB) 0) }
    if ($Byte -ge 1KB) { return '{0} KB' -f (Format-HDNumero ($Byte / 1KB) 0) }
    '{0} byte' -f [int64]$Byte
}

function Format-HDData {
    # Formati numerici con la cultura invariante: "/" e ":" restano tali anche se
    # le impostazioni regionali del PC usano altri separatori.
    param([datetime]$Data, [string]$Formato = 'dd/MM/yyyy HH:mm')
    $Data.ToString($Formato, $script:HDInvariante)
}

function Format-HDTempoTrascorso {
    # TimeSpan -> "3 giorni e 4 ore"
    param([timespan]$Intervallo)
    $g = [int][math]::Floor($Intervallo.TotalDays)
    $o = $Intervallo.Hours
    $parti = @()
    if ($g -eq 1) { $parti += '1 giorno' } elseif ($g -gt 1) { $parti += "$g giorni" }
    if ($o -eq 1) { $parti += '1 ora' } elseif ($o -gt 1) { $parti += "$o ore" }
    if ($parti.Count -eq 0) { return "$([int]$Intervallo.TotalMinutes) minuti" }
    $parti -join ' e '
}

function ConvertTo-HDValoreTesto {
    # Valori pronti per CSV/HTML in formato italiano (12,5 - 01/11/2026 - Sì/No)
    param($Valore)
    if ($null -eq $Valore) { return '' }
    if ($Valore -is [datetime]) {
        if ($Valore.TimeOfDay -eq [timespan]::Zero) { return $Valore.ToString('dd/MM/yyyy', $script:HDInvariante) }
        return $Valore.ToString('dd/MM/yyyy HH:mm', $script:HDInvariante)
    }
    if ($Valore -is [bool]) { if ($Valore) { return 'Sì' } else { return 'No' } }
    if ($Valore -is [double] -or $Valore -is [single] -or $Valore -is [decimal]) {
        return $Valore.ToString('0.##', $script:HDCultura)
    }
    if ($Valore -is [System.Collections.IEnumerable] -and $Valore -isnot [string]) {
        return (@($Valore | ForEach-Object { ConvertTo-HDValoreTesto $_ }) -join ', ')
    }
    [string]$Valore
}

function ConvertTo-HDHtml {
    # Escape HTML: il testo delle note non deve mai diventare codice nella pagina
    param([string]$Testo)
    if ($null -eq $Testo) { return '' }
    [System.Net.WebUtility]::HtmlEncode($Testo)
}

#endregion

#region File di testo e CSV

function Get-HDCodificaFile {
    # Riconosce la codifica: BOM UTF-8/UTF-16, UTF-8 senza BOM (Blocco note recente)
    # oppure ANSI/Windows-1252 (Blocco note vecchio, Windows PowerShell 5.1).
    param([byte[]]$Byte)
    if ($Byte.Length -ge 3 -and $Byte[0] -eq 0xEF -and $Byte[1] -eq 0xBB -and $Byte[2] -eq 0xBF) {
        return @{ Codifica = (New-Object System.Text.UTF8Encoding($false)); Salto = 3 }
    }
    if ($Byte.Length -ge 2 -and $Byte[0] -eq 0xFF -and $Byte[1] -eq 0xFE) {
        return @{ Codifica = [System.Text.Encoding]::Unicode; Salto = 2 }
    }
    if ($Byte.Length -ge 2 -and $Byte[0] -eq 0xFE -and $Byte[1] -eq 0xFF) {
        return @{ Codifica = [System.Text.Encoding]::BigEndianUnicode; Salto = 2 }
    }
    try {
        $rigoroso = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$rigoroso.GetString($Byte)
        return @{ Codifica = (New-Object System.Text.UTF8Encoding($false)); Salto = 0 }
    }
    catch {
        try { $ansi = [System.Text.Encoding]::GetEncoding(1252) }
        catch { $ansi = [System.Text.Encoding]::GetEncoding(28591) }
        return @{ Codifica = $ansi; Salto = 0 }
    }
}

function Read-HDFileTesto {
    param([Parameter(Mandatory)][string]$Percorso)
    $p = Resolve-HDPercorso $Percorso
    $byte = [System.IO.File]::ReadAllBytes($p)
    $info = Get-HDCodificaFile -Byte $byte
    $info.Codifica.GetString($byte, $info.Salto, $byte.Length - $info.Salto)
}

function Write-HDFileTesto {
    # Scrive in UTF-8 con BOM: lo leggono bene Excel, Blocco note e PowerShell 5.1.
    param(
        [Parameter(Mandatory)][string]$Percorso,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Testo,
        [switch]$Accoda
    )
    $p = Resolve-HDPercorso $Percorso
    $cartella = Split-Path -Parent $p
    if ($cartella) { [void][System.IO.Directory]::CreateDirectory($cartella) }
    $codifica = New-Object System.Text.UTF8Encoding($true)
    if ($Accoda) { [System.IO.File]::AppendAllText($p, $Testo, $codifica) }
    else { [System.IO.File]::WriteAllText($p, $Testo, $codifica) }
}

function ConvertTo-HDRigaTesto {
    # Oggetto -> oggetto con soli valori testuali (numeri e date all'italiana)
    param([Parameter(Mandatory)]$Oggetto, [string[]]$Colonne)
    if (-not $Colonne) { $Colonne = @($Oggetto.PSObject.Properties | ForEach-Object { $_.Name }) }
    $riga = [ordered]@{}
    foreach ($c in $Colonne) {
        $prop = $Oggetto.PSObject.Properties[$c]
        if ($prop) { $riga[$c] = ConvertTo-HDValoreTesto $prop.Value } else { $riga[$c] = '' }
    }
    [pscustomobject]$riga
}

#endregion

#region HTML

$script:HDCss = @'
:root {
  --inchiostro: #18212b; --testo: #2f3b47; --tenue: #6b7785; --filo: #e3e7ec; --fondo: #ffffff; --carta: #f6f7f9;
  --critico: #b42318; --attenzione: #b45309; --errore: #7a3cae; --nonverificato: #8a94a0; --info: #2f6fb3; --ok: #2e7d4f;
}
* { box-sizing: border-box; }
html { -webkit-text-size-adjust: 100%; }
body { margin: 0; background: var(--carta); color: var(--testo);
       font-family: "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
       font-size: 14px; line-height: 1.5; }
.pagina { max-width: 860px; margin: 32px auto; background: var(--fondo); padding: 44px 52px 32px;
          border: 1px solid var(--filo); border-radius: 4px; }
.etichetta { font-size: 11px; letter-spacing: .12em; text-transform: uppercase; color: var(--tenue); font-weight: 600; }
h1 { font-size: 30px; line-height: 1.15; margin: 6px 0 4px; color: var(--inchiostro); font-weight: 600; letter-spacing: -.01em; }
.sottotitolo { color: var(--tenue); }
.testata { display: flex; justify-content: space-between; align-items: flex-end; gap: 24px;
           padding-bottom: 22px; border-bottom: 1px solid var(--filo); }
.testata .quando { text-align: right; color: var(--tenue); white-space: nowrap; font-variant-numeric: tabular-nums; }
.testata .quando strong { display: block; color: var(--inchiostro); font-weight: 600; }
h2 { font-size: 11px; letter-spacing: .12em; text-transform: uppercase; color: var(--tenue); font-weight: 600;
     margin: 34px 0 12px; }
.esito { margin-top: 26px; }
.esito .frase { font-size: 19px; color: var(--inchiostro); font-weight: 600; }
.barra { display: flex; height: 6px; border-radius: 3px; overflow: hidden; margin: 12px 0 10px; background: var(--filo); }
.barra span { display: block; height: 100%; }
.legenda { display: flex; flex-wrap: wrap; gap: 6px 18px; margin: 0; padding: 0; list-style: none; color: var(--tenue); font-size: 13px; }
.legenda li { display: flex; align-items: center; gap: 7px; }
.punto { width: 8px; height: 8px; border-radius: 50%; display: inline-block; flex: none; }
.c-critico { background: var(--critico); } .c-attenzione { background: var(--attenzione); } .c-errore { background: var(--errore); }
.c-nonverificato { background: var(--nonverificato); } .c-info { background: var(--info); } .c-ok { background: var(--ok); }
.t-critico { color: var(--critico); } .t-attenzione { color: var(--attenzione); } .t-errore { color: var(--errore); }
.t-nonverificato { color: var(--nonverificato); } .t-info { color: var(--info); } .t-ok { color: var(--ok); }
.dati { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0 40px; margin: 0; }
.dati div { display: flex; justify-content: space-between; gap: 16px; padding: 7px 0; border-bottom: 1px solid var(--filo); }
.dati dt { color: var(--tenue); margin: 0; }
.dati dd { margin: 0; text-align: right; color: var(--inchiostro); overflow-wrap: anywhere; }
.voce { border-left: 3px solid var(--filo); padding: 2px 0 2px 16px; margin: 0 0 18px; break-inside: avoid; }
.voce.v-critico { border-color: var(--critico); } .voce.v-attenzione { border-color: var(--attenzione); } .voce.v-errore { border-color: var(--errore); }
.voce .testa { display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap; }
.voce .nome { font-weight: 600; color: var(--inchiostro); font-size: 15px; }
.voce .stato { font-size: 11px; letter-spacing: .08em; text-transform: uppercase; font-weight: 700; }
.voce .area { color: var(--tenue); font-size: 13px; margin-left: auto; }
.voce p { margin: 4px 0 0; }
.voce .fare { color: var(--testo); }
.voce .fare b { color: var(--inchiostro); font-weight: 600; }
table { width: 100%; border-collapse: collapse; }
th { text-align: left; font-size: 12px; font-weight: 600; color: var(--tenue); padding: 0 10px 8px 0; border-bottom: 1px solid var(--filo); }
td { text-align: left; vertical-align: top; padding: 8px 10px 8px 0; border-bottom: 1px solid var(--filo); }
td.controllo { white-space: nowrap; color: var(--inchiostro); }
td.controllo .punto { margin-right: 8px; position: relative; top: -1px; }
td.area { color: var(--tenue); white-space: nowrap; }
td .nota { color: var(--tenue); font-size: 13px; margin-top: 2px; }
.piede { margin-top: 36px; padding-top: 14px; border-top: 1px solid var(--filo); color: var(--tenue); font-size: 12px;
         display: flex; justify-content: space-between; gap: 16px; flex-wrap: wrap; }
.azioni { max-width: 860px; margin: 20px auto -16px; text-align: right; }
.azioni button { font: inherit; font-size: 13px; padding: 6px 14px; color: var(--inchiostro); background: var(--fondo);
                 border: 1px solid var(--filo); border-radius: 4px; cursor: pointer; }
.azioni button:hover { border-color: var(--tenue); }
@media (max-width: 700px) {
  .pagina { margin: 0; padding: 24px 18px; border: 0; border-radius: 0; }
  .testata { flex-direction: column; align-items: flex-start; } .testata .quando { text-align: left; }
  .dati { grid-template-columns: 1fr; } td.area { display: none; } td.controllo { white-space: normal; }
}
@media print {
  body { background: #fff; font-size: 11.5px; } .pagina { margin: 0; padding: 0; border: 0; max-width: none; }
  .no-print { display: none !important; } tr, .voce { break-inside: avoid; }
  .barra span, .punto, .voce { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  @page { size: A4; margin: 16mm 15mm; }
}
'@

function Get-HDClasseStato {
    param([string]$Stato)
    's-' + ((ConvertTo-HDChiave $Stato) -replace '[^a-z]', '')
}

function ConvertTo-HDPaginaHtml {
    param([Parameter(Mandatory)][string]$Titolo, [Parameter(Mandatory)][string]$Corpo)
    $t = ConvertTo-HDHtml $Titolo
    @"
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$t</title>
<style>
$script:HDCss
</style>
</head>
<body>
$Corpo
</body>
</html>
"@
}

#endregion

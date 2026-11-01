# Funzioni comuni del toolkit (private: non esportate dal modulo).
# Regola d'oro: compatibili con Windows PowerShell 5.1 (niente ??, ?:, &&, Join-String, ecc.).

$script:HDCultura = [System.Globalization.CultureInfo]::GetCultureInfo('it-IT')
$script:HDInvariante = [System.Globalization.CultureInfo]::InvariantCulture

#region Configurazione e cartelle

function Get-HDConfigPredefinita {
    @{
        Azienda         = 'La tua azienda'
        Tecnico         = ''
        CartellaReport  = ''
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
    if (-not $cfg.Tecnico) { $cfg.Tecnico = [Environment]::UserName }
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
:root { --testo:#1f2933; --tenue:#52606d; --bordo:#d9e2ec; --sfondo:#ffffff; --riquadro:#f5f7fa;
        --ok:#1e7e34; --att:#9a6700; --crit:#b42318; --info:#1d5fa8; --nv:#6b7280; --err:#7a2e9a; }
* { box-sizing: border-box; }
body { font-family: "Segoe UI", Arial, sans-serif; color: var(--testo); background: var(--sfondo);
       margin: 0; padding: 24px; font-size: 14px; line-height: 1.45; }
.pagina { max-width: 900px; margin: 0 auto; }
h1 { font-size: 22px; margin: 0 0 4px; }
h2 { font-size: 13px; text-transform: uppercase; letter-spacing: .05em; color: var(--tenue);
     border-bottom: 1px solid var(--bordo); padding-bottom: 4px; margin: 24px 0 10px; }
table { width: 100%; border-collapse: collapse; }
th, td { text-align: left; vertical-align: top; padding: 6px 8px; border-bottom: 1px solid var(--bordo); }
th { background: var(--riquadro); font-weight: 600; }
td.ora { white-space: nowrap; width: 1%; font-variant-numeric: tabular-nums; }
.badge { display: inline-block; padding: 1px 9px; border-radius: 10px; font-size: 12px; font-weight: 600;
         color: #fff; white-space: nowrap; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.s-ok { background: var(--ok); } .s-attenzione { background: var(--att); } .s-critico { background: var(--crit); }
.s-info { background: var(--info); } .s-nonverificato { background: var(--nv); } .s-errore { background: var(--err); }
.tenue { color: var(--tenue); }
.intestazione { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start;
                border-bottom: 2px solid var(--testo); padding-bottom: 12px; margin-bottom: 8px; }
.intestazione .destra { text-align: right; }
.griglia { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0 28px; }
.campo { display: flex; gap: 10px; padding: 5px 0; border-bottom: 1px dotted var(--bordo); }
.campo .etichetta { min-width: 110px; color: var(--tenue); }
.avvisi { background: #fff8e1; border: 1px solid #f0c36d; padding: 10px 14px; border-radius: 6px; margin: 12px 0; }
.avvisi ul { margin: 6px 0 0; padding-left: 20px; }
.riepilogo td:first-child { color: var(--tenue); width: 35%; }
.firme { display: grid; grid-template-columns: 1fr 1fr; gap: 48px; margin-top: 56px; }
.firma { border-top: 1px solid var(--testo); padding-top: 6px; text-align: center; color: var(--tenue); }
.piede { margin-top: 32px; font-size: 11px; color: var(--tenue); text-align: center; }
.azioni { margin: 0 0 16px; text-align: right; }
.azioni button { font: inherit; padding: 6px 14px; border: 1px solid var(--bordo); background: var(--riquadro);
                 border-radius: 6px; cursor: pointer; }
@media (max-width: 640px) { body { padding: 16px; } .griglia, .firme { grid-template-columns: 1fr; }
                            .intestazione { flex-direction: column; } .intestazione .destra { text-align: left; } }
@media print { body { padding: 0; font-size: 12px; } .no-print { display: none !important; }
               tr, .firme, .avvisi { page-break-inside: avoid; } @page { size: A4; margin: 15mm; } }
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
<div class="pagina">
$Corpo
</div>
</body>
</html>
"@
}

#endregion

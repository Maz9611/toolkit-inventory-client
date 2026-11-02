#Requires -Version 5.1
<#
.SYNOPSIS
Finestra del Toolkit Help Desk (si apre con Avvia-Toolkit.cmd oppure con ToolkitHelpDesk.exe).

.PARAMETER Verifica
Costruisce la finestra e si chiude subito senza mostrarla: la usa la prova automatica.

.NOTES
Dentro ToolkitHelpDesk.exe questo script gira senza file su disco: il modulo è già caricato,
$PSScriptRoot è vuoto e $HDPercorsoExe contiene il percorso dell'exe.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Funzioni interne della finestra, non comandi')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleTypes', '', Justification = 'Le librerie Windows Forms vengono caricate con Add-Type all''avvio')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Comandi del modulo HelpDeskToolkit caricato all''avvio')]
[CmdletBinding()]
param([switch]$Verifica)

$ErrorActionPreference = 'Stop'
$script:FileErrori = Join-Path ([System.IO.Path]::GetTempPath()) 'HelpDeskToolkit-errori.txt'

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
}
catch {
    # Niente finestra su questo PC: si apre il menu testuale
    if (-not $Verifica -and $PSScriptRoot -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Toolkit.ps1'))) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f (Join-Path $PSScriptRoot 'Toolkit.ps1')))
    }
    exit 2
}

#region Aiuti per costruire la finestra

function Show-Errore {
    param([string]$Messaggio)
    try { Add-Content -LiteralPath $script:FileErrori -Value ('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Messaggio) } catch { Write-Verbose 'Log errori non scrivibile' }
    if ($script:Stato) { $script:Stato.Text = 'Errore: ' + $Messaggio }
    [void][System.Windows.Forms.MessageBox]::Show($Messaggio, 'Toolkit Help Desk', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}

function Show-Info {
    param([string]$Messaggio)
    [void][System.Windows.Forms.MessageBox]::Show($Messaggio, 'Toolkit Help Desk', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function New-Etichetta {
    param([string]$Testo, [switch]$Grassetto)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Testo
    $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 2)
    if ($Grassetto) { $l.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10) }
    $l
}

function New-Pulsante {
    param([string]$Testo, [int]$Larghezza = 0)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Testo
    $b.AutoSize = $true
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)
    if ($Larghezza -gt 0) { $b.AutoSize = $false; $b.Width = $Larghezza; $b.Height = 34 }
    $b
}

function New-Casella {
    param([int]$Larghezza = 340, [switch]$Multilinea)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $Larghezza
    if ($Multilinea) { $t.Multiline = $true; $t.Height = 90; $t.ScrollBars = 'Vertical' }
    $t
}

function New-Griglia {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.AutoSizeColumnsMode = 'Fill'
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.BorderStyle = 'None'
    $g
}

function Set-Colonne {
    param($Griglia, [string[]]$Titoli)
    $Griglia.Rows.Clear()
    $Griglia.Columns.Clear()
    foreach ($t in $Titoli) { [void]$Griglia.Columns.Add(($t -replace '\W', ''), $t) }
}

function Set-Occupato {
    param([bool]$Occupato, [string]$Testo)
    $script:Finestra.UseWaitCursor = $Occupato
    foreach ($b in $script:Pulsanti) { $b.Enabled = -not $Occupato }
    if ($Testo) { $script:Stato.Text = $Testo }
    [System.Windows.Forms.Application]::DoEvents()
}

#endregion

#region Censimento: file scelto e dati del PC

function Update-FileCensimento {
    param([string]$Percorso)
    $script:FileCensimento = $Percorso
    if ($Percorso) { $script:CasellaFile.Text = $Percorso }
    else { $script:CasellaFile.Text = '(nessun censimento: crealo con "Nuovo..." oppure con una scansione)' }
}

function Confirm-FileCensimento {
    # Se non c'è ancora un Excel lo crea chiedendo il nome del cliente; $false se annullato
    if ($script:FileCensimento) { return $true }
    $cliente = [Microsoft.VisualBasic.Interaction]::InputBox('Non c''è ancora un censimento. Nome del cliente per il nuovo Excel:', 'Nuovo censimento', '')
    if (-not $cliente.Trim()) { return $false }
    $f = New-Censimento -Cliente $cliente.Trim() -Confirm:$false
    Update-FileCensimento $f.FullName
    $true
}

function Get-TabellaAttuale {
    if (-not $script:FileCensimento) { return $null }
    & $script:Modulo { param($p) Get-HDTabellaCensimento $p } $script:FileCensimento
}

function Show-DatiPc {
    Set-Occupato $true 'Lettura dei dati di questo PC (10-20 secondi)...'
    try {
        $r = Get-InventarioPC
        if (-not $r) { throw 'Non sono riuscito a leggere i dati di questo PC.' }
        $script:Record = $r
        $seriale = $r.SerialNumber
        if (-not $seriale) { $seriale = '(non presente nel BIOS)' }
        $ip = $r.IndirizzoIP
        if ($ip) { $ip = '{0} ({1})' -f $ip, $r.TipoIP } else { $ip = '(non collegato alla rete)' }
        $cynet = $r.Cynet
        if ($r.Cynet -eq 'Sì' -and -not $r.CynetInEsecuzione) { $cynet = 'Sì (installato, ma l''agente non è in esecuzione)' }
        Set-Colonne $script:GrigliaPc @('Campo', 'Valore')
        $script:GrigliaPc.Columns[0].FillWeight = 35
        $voci = @(
            @('Hostname', $r.Hostname), @('Indirizzo IP', $ip), @('MAC Address', $r.MAC), @('Device Type', $r.DeviceType),
            @('Manufacturer', $r.Manufacturer), @('Model', $r.Model), @('Nome Utente', ('{0} ({1})' -f $r.NomeUtente, $r.FonteNomeUtente)),
            @('Dominio', $r.Dominio), @('Serial Number', $seriale), @('Operating System', $r.OS), @('Operating System Version', $r.OSVersion),
            @('Memory', $r.Memory), @('Processor Type', $r.Processor), @('Free Disk Space %', $r.FreeDiskPct), @('Logical Disk Size', $r.DiskSize),
            @('Cynet', $cynet), @('BitLocker', $r.BitLocker)
        )
        foreach ($v in $voci) { [void]$script:GrigliaPc.Rows.Add([string]$v[0], [string]$v[1]) }
        # Dati del foglio Dettagli (Windows 11, dischi, monitor, software...)
        if ($r.Dettagli) {
            $colonneDettagli = & $script:Modulo { $script:HDColonneDettagli }
            foreach ($c in $colonneDettagli) {
                if ($c.Chiave -in 'Hostname', 'SerialNumber', 'Rilevato' -or -not [string]$r.Dettagli[$c.Chiave]) { continue }
                $i = $script:GrigliaPc.Rows.Add([string]$c.Titolo, [string]$r.Dettagli[$c.Chiave])
                $valore = [string]$r.Dettagli[$c.Chiave]
                if (($c.Chiave -eq 'Windows11' -and $valore -match '^No') -or $valore -match 'GUASTO|^Non attivato') { $script:GrigliaPc.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose }
                elseif ($c.Chiave -eq 'Windows11' -and $valore -match '^Da verificare') { $script:GrigliaPc.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon }
            }
        }
        $attuali = @{}
        $tabella = Get-TabellaAttuale
        if ($tabella) {
            $trovata = & $script:Modulo { param($t, $x) Find-HDRigaCensimento -Righe $t.Righe -Mac (@($x.TuttiMAC) + @($x.MAC)) -Seriale $x.SerialNumber -Hostname $x.Hostname } $tabella $r
            if ($trovata) {
                $script:EtichettaRiga.Text = 'Già nel censimento alla riga {0} ({1}): verrà aggiornato.' -f $trovata.Riga.Numero, $trovata.Motivo
                $attuali = $trovata.Riga.Valori
            }
            else { $script:EtichettaRiga.Text = 'Non è nel censimento: verrà aggiunto in fondo.' }
        }
        else { $script:EtichettaRiga.Text = 'Nessun Excel di censimento scelto: verrà creato al salvataggio.' }
        $script:CasellaUtilizzatore.Text = [string]$attuali['Utilizzatore']
        $script:CasellaLocazione.Text = [string]$attuali['Locazione']
        $script:CasellaNote.Text = [string]$attuali['Note']
        $script:CasellaNomeUtente.Text = $r.NomeUtente
        if ($attuali['NomeUtente']) { $script:CasellaNomeUtente.Text = [string]$attuali['NomeUtente'] }
        $script:Stato.Text = 'Dati del PC letti. Completa i campi a destra e premi "Salva nell''Excel".'
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false }
}

function Save-DatiPc {
    if (-not $script:Record) { Show-Info 'Prima leggi i dati del PC.'; return }
    if (-not (Confirm-FileCensimento)) { return }
    Set-Occupato $true 'Scrittura nell''Excel...'
    try {
        $p = @{ Censimento = $script:FileCensimento; Record = $script:Record; NonChiedere = $true; Confirm = $false }
        if ($script:CasellaUtilizzatore.Text.Trim()) { $p.Utilizzatore = $script:CasellaUtilizzatore.Text.Trim() }
        if ($script:CasellaLocazione.Text.Trim()) { $p.Locazione = $script:CasellaLocazione.Text.Trim() }
        if ($script:CasellaNote.Text.Trim()) { $p.Note = $script:CasellaNote.Text.Trim() }
        if ($script:CasellaNomeUtente.Text.Trim()) { $p.NomeUtente = $script:CasellaNomeUtente.Text.Trim() }
        $esito = Add-PCCensimento @p
        if (-not $esito.Scritto) { throw ("Excel non aggiornato: {0}`nI dati sono al sicuro e verranno inseriti al prossimo salvataggio." -f $esito.Errore) }
        $e = @($esito.Esiti)[0]
        if ($e.Azione -eq 'Aggiunto') { $testo = 'Fatto: {0} aggiunto alla riga {1}.' -f $e.Hostname, $e.Riga }
        else { $testo = 'Fatto: riga {0} aggiornata ({1}).' -f $e.Riga, $e.Motivo }
        if ($script:Record.Dettagli) { $testo += ' Dettagli e programmi sono nei fogli "Dettagli" e "Software".' }
        if ($script:Record.Dettagli -and -not $script:Record.Dettagli['Smart']) { $testo += ' Per leggere lo SMART dei dischi avvia il programma come amministratore.' }
        $script:Stato.Text = $testo
        $script:EtichettaRiga.Text = $testo
        Show-Info $testo
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false }
}

#endregion

#region Scansione della rete

function Update-ColonnaNelCensimento {
    $tabella = Get-TabellaAttuale
    foreach ($riga in $script:GrigliaRete.Rows) {
        $valore = ''
        if ($tabella) {
            $t = & $script:Modulo { param($tab, $mac, $nome, $ip) Find-HDRigaCensimento -Righe $tab.Righe -Mac @($mac) -Hostname $nome -IP $ip -ConsentiIP } $tabella ([string]$riga.Cells[2].Value) ([string]$riga.Cells[1].Value) ([string]$riga.Cells[0].Value)
            if ($t) { $valore = 'riga {0}' -f $t.Riga.Numero }
        }
        $riga.Cells[5].Value = $valore
    }
}

function Start-Scansione {
    Set-Occupato $true 'Scansione della rete in corso...'
    try {
        $p = @{ Confirm = $false; Avanzamento = { param($testo) $script:EtichettaScan.Text = $testo; [System.Windows.Forms.Application]::DoEvents() } }
        if ($script:CasellaRete.Text.Trim()) { $p.Rete = $script:CasellaRete.Text.Trim() }
        $s = Invoke-ScansioneRete @p
        $script:Scansione = $s
        Set-Colonne $script:GrigliaRete @('IP', 'Nome', 'MAC', 'Gruppo', 'Trovato con', 'Nel censimento')
        foreach ($d in $s.Dispositivi) { [void]$script:GrigliaRete.Rows.Add($d.IP, $d.Hostname, $d.MAC, $d.Gruppo, $d.Fonte, '') }
        Update-ColonnaNelCensimento
        $script:PulsanteAggiungi.Enabled = ($s.Dispositivi.Count -gt 0)
        $script:EtichettaScan.Text = 'Trovati {0} dispositivi in {1} secondi sulla rete {2}.' -f $s.Dispositivi.Count, $s.Secondi, $s.Rete
        $script:Stato.Text = $script:EtichettaScan.Text
    }
    catch { Show-Errore $_.Exception.Message }
    finally {
        Set-Occupato $false
        $script:PulsanteAggiungi.Enabled = [bool]($script:Scansione -and @($script:Scansione.Dispositivi).Count -gt 0)
    }
}

function Show-EsitoImport {
    param($Esito)
    $testo = 'Excel aggiornato: {0} dispositivi aggiunti, {1} aggiornati, {2} invariati.' -f $Esito.Aggiunti, $Esito.Aggiornati, $Esito.Invariati
    $script:Stato.Text = $testo
    Show-Info $testo
}

function Add-ScansioneAlCensimento {
    if (-not $script:Scansione -or @($script:Scansione.Dispositivi).Count -eq 0) { return }
    if (-not (Confirm-FileCensimento)) { return }
    Set-Occupato $true 'Scrittura dei dispositivi nell''Excel...'
    try {
        $esito = & $script:Modulo { param($d, $f, $r) Import-HDDispositiviNelCensimento -Dispositivi $d -Formato ('scansione integrata di ' + $r) -Censimento $f -NonChiedere } @($script:Scansione.Dispositivi) $script:FileCensimento $script:Scansione.Rete
        Update-ColonnaNelCensimento
        Show-EsitoImport $esito
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false; $script:PulsanteAggiungi.Enabled = $true }
}

function Import-FileScanner {
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Title = 'Export di Advanced IP Scanner'
    $d.Filter = 'Export di Advanced IP Scanner (*.csv;*.xml)|*.csv;*.xml|Tutti i file (*.*)|*.*'
    if ($d.ShowDialog($script:Finestra) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    if (-not (Confirm-FileCensimento)) { return }
    Set-Occupato $true 'Import del file di Advanced IP Scanner...'
    try {
        $esito = Import-ScansioneRete -Percorso $d.FileName -Censimento $script:FileCensimento -NonChiedere -Confirm:$false
        Show-EsitoImport $esito
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false }
}

#endregion

#region Verifiche

function Start-Verifica {
    Set-Occupato $true 'Verifica della postazione in corso (pochi secondi)...'
    try {
        & $script:Modulo { $script:HDUltimaVerifica = $null }
        $esiti = @(Test-Postazione -PassThru -Silenzioso)
        $script:UltimoReport = & $script:Modulo { $script:HDUltimaVerifica }
        Set-Colonne $script:GrigliaVerifiche @('Stato', 'Controllo', 'Dettaglio', 'Suggerimento')
        $script:GrigliaVerifiche.Columns[0].FillWeight = 18
        $script:GrigliaVerifiche.Columns[1].FillWeight = 25
        $colori = @{ 'OK' = 'Honeydew'; 'ATTENZIONE' = 'LemonChiffon'; 'CRITICO' = 'MistyRose'; 'ERRORE' = 'Thistle' }
        foreach ($e in $esiti) {
            $i = $script:GrigliaVerifiche.Rows.Add($e.Stato, $e.Controllo, $e.Dettaglio, $e.Suggerimento)
            if ($colori.ContainsKey($e.Stato)) { $script:GrigliaVerifiche.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromName($colori[$e.Stato]) }
        }
        $script:PulsanteReport.Enabled = [bool]$script:UltimoReport
        $conteggio = ($esiti | Group-Object Stato | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name }) -join ', '
        $script:Stato.Text = 'Verifica fatta: ' + $conteggio
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false; $script:PulsanteReport.Enabled = [bool]$script:UltimoReport }
}

function Show-SpazioDisco {
    Set-Occupato $true 'Lettura dello spazio su disco...'
    try {
        Set-Colonne $script:GrigliaVerifiche @('Unità', 'Libero (GB)', 'Totale (GB)', 'Libero %', 'Stato')
        $colori = @{ 'ATTENZIONE' = 'LemonChiffon'; 'CRITICO' = 'MistyRose' }
        foreach ($v in @(Get-SpazioDisco)) {
            $i = $script:GrigliaVerifiche.Rows.Add($v.Unita, $v.LiberoGB, $v.TotaleGB, $v.LiberoPct, $v.Stato)
            if ($colori.ContainsKey($v.Stato)) { $script:GrigliaVerifiche.Rows[$i].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromName($colori[$v.Stato]) }
        }
        $script:Stato.Text = 'Spazio su disco letto.'
    }
    catch { Show-Errore $_.Exception.Message }
    finally { Set-Occupato $false }
}

#endregion

try {
    if (-not (Get-Module HelpDeskToolkit)) {
        Import-Module (Join-Path (Join-Path $PSScriptRoot 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
    }
    $script:Modulo = Get-Module HelpDeskToolkit
    $script:Record = $null
    $script:Scansione = $null
    $script:UltimoReport = $null
    $script:FileCensimento = $null

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $finestra = New-Object System.Windows.Forms.Form
    $finestra.Text = 'Toolkit Help Desk'
    $finestra.Size = New-Object System.Drawing.Size(1100, 740)
    $finestra.MinimumSize = New-Object System.Drawing.Size(900, 600)
    $finestra.StartPosition = 'CenterScreen'
    $finestra.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $exe = Get-Variable -Name HDPercorsoExe -ValueOnly -ErrorAction SilentlyContinue
    if ($exe) {
        try { $finestra.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exe) } catch { Write-Verbose 'Icona non disponibile' }
    }
    $script:Finestra = $finestra

    # Barra in alto: quale Excel di censimento
    $alto = New-Object System.Windows.Forms.TableLayoutPanel
    $alto.Dock = 'Top'
    $alto.Height = 52
    $alto.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $alto.ColumnCount = 5
    [void]$alto.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$alto.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    1..3 | ForEach-Object { [void]$alto.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) }
    $script:CasellaFile = New-Object System.Windows.Forms.TextBox
    $script:CasellaFile.ReadOnly = $true
    $script:CasellaFile.Dock = 'Fill'
    $pulsanteScegli = New-Pulsante 'Scegli...'
    $pulsanteNuovo = New-Pulsante 'Nuovo...'
    $pulsanteApri = New-Pulsante 'Apri in Excel'
    $alto.Controls.Add((New-Etichetta 'Excel del censimento:' -Grassetto), 0, 0)
    $alto.Controls.Add($script:CasellaFile, 1, 0)
    $alto.Controls.Add($pulsanteScegli, 2, 0)
    $alto.Controls.Add($pulsanteNuovo, 3, 0)
    $alto.Controls.Add($pulsanteApri, 4, 0)

    $schede = New-Object System.Windows.Forms.TabControl
    $schede.Dock = 'Fill'
    $schede.Padding = New-Object System.Drawing.Point(14, 6)

    # Scheda 1: censisci questo PC
    $schedaPc = New-Object System.Windows.Forms.TabPage
    $schedaPc.Text = 'Censisci questo PC'
    $schedaPc.Padding = New-Object System.Windows.Forms.Padding(8)
    $grigliaPcContenitore = New-Object System.Windows.Forms.TableLayoutPanel
    $grigliaPcContenitore.Dock = 'Fill'
    $grigliaPcContenitore.ColumnCount = 2
    [void]$grigliaPcContenitore.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    [void]$grigliaPcContenitore.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $script:GrigliaPc = New-Griglia
    $destra = New-Object System.Windows.Forms.FlowLayoutPanel
    $destra.Dock = 'Fill'
    $destra.FlowDirection = 'TopDown'
    $destra.WrapContents = $false
    $destra.AutoScroll = $true
    $destra.Padding = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
    $script:EtichettaRiga = New-Etichetta 'Premi "Leggi i dati del PC".' -Grassetto
    $script:EtichettaRiga.MaximumSize = New-Object System.Drawing.Size(360, 0)
    $script:CasellaUtilizzatore = New-Casella
    $script:CasellaLocazione = New-Casella
    $script:CasellaNote = New-Casella -Multilinea
    $script:CasellaNomeUtente = New-Casella
    $pulsanteLeggi = New-Pulsante 'Leggi i dati del PC' 340
    $pulsanteSalva = New-Pulsante 'Salva nell''Excel' 340
    $pulsanteSalva.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $aiuto = New-Etichetta 'Lascia vuoto un campo per non cambiarlo. Per cancellare un valore fallo direttamente nell''Excel.'
    $aiuto.MaximumSize = New-Object System.Drawing.Size(340, 0)
    $aiuto.ForeColor = [System.Drawing.SystemColors]::GrayText
    foreach ($c in @($script:EtichettaRiga, (New-Etichetta 'Utilizzatore'), $script:CasellaUtilizzatore, (New-Etichetta 'Locazione'), $script:CasellaLocazione,
            (New-Etichetta 'Note'), $script:CasellaNote, (New-Etichetta 'Nome Utente (correggi se non è giusto)'), $script:CasellaNomeUtente,
            $pulsanteLeggi, $pulsanteSalva, $aiuto)) { $destra.Controls.Add($c) }
    $grigliaPcContenitore.Controls.Add($script:GrigliaPc, 0, 0)
    $grigliaPcContenitore.Controls.Add($destra, 1, 0)
    $schedaPc.Controls.Add($grigliaPcContenitore)

    # Scheda 2: scansione della rete
    $schedaRete = New-Object System.Windows.Forms.TabPage
    $schedaRete.Text = 'Scansione della rete'
    $schedaRete.Padding = New-Object System.Windows.Forms.Padding(8)
    $layoutRete = New-Object System.Windows.Forms.TableLayoutPanel
    $layoutRete.Dock = 'Fill'
    $layoutRete.RowCount = 3
    [void]$layoutRete.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layoutRete.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layoutRete.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $comandiRete = New-Object System.Windows.Forms.FlowLayoutPanel
    $comandiRete.AutoSize = $true
    $comandiRete.Dock = 'Fill'
    $script:CasellaRete = New-Casella -Larghezza 340
    $pulsanteScansiona = New-Pulsante 'Scansiona la rete'
    $script:PulsanteAggiungi = New-Pulsante 'Aggiungi all''Excel'
    $script:PulsanteAggiungi.Enabled = $false
    $pulsanteFileScanner = New-Pulsante 'Importa file di Advanced IP Scanner...'
    foreach ($c in @((New-Etichetta 'Rete:'), $script:CasellaRete, $pulsanteScansiona, $script:PulsanteAggiungi, $pulsanteFileScanner)) { $comandiRete.Controls.Add($c) }
    $script:EtichettaScan = New-Etichetta 'Di solito la rete proposta è già quella giusta: premi "Scansiona la rete" (10-20 secondi per una /24). Per più reti separale con una virgola: 192.168.1.0/24, 192.168.2.0/24'
    $script:GrigliaRete = New-Griglia
    Set-Colonne $script:GrigliaRete @('IP', 'Nome', 'MAC', 'Gruppo', 'Trovato con', 'Nel censimento')
    $layoutRete.Controls.Add($comandiRete, 0, 0)
    $layoutRete.Controls.Add($script:EtichettaScan, 0, 1)
    $layoutRete.Controls.Add($script:GrigliaRete, 0, 2)
    $schedaRete.Controls.Add($layoutRete)

    # Scheda 3: verifiche
    $schedaVerifiche = New-Object System.Windows.Forms.TabPage
    $schedaVerifiche.Text = 'Verifica postazione e dischi'
    $schedaVerifiche.Padding = New-Object System.Windows.Forms.Padding(8)
    $layoutVerifiche = New-Object System.Windows.Forms.TableLayoutPanel
    $layoutVerifiche.Dock = 'Fill'
    $layoutVerifiche.RowCount = 2
    [void]$layoutVerifiche.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layoutVerifiche.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $comandiVerifiche = New-Object System.Windows.Forms.FlowLayoutPanel
    $comandiVerifiche.AutoSize = $true
    $comandiVerifiche.Dock = 'Fill'
    $pulsanteVerifica = New-Pulsante 'Verifica rapida della postazione'
    $script:PulsanteReport = New-Pulsante 'Apri il report'
    $script:PulsanteReport.Enabled = $false
    $pulsanteSpazio = New-Pulsante 'Spazio su disco'
    $pulsanteCartellaReport = New-Pulsante 'Cartella dei report'
    foreach ($c in @($pulsanteVerifica, $script:PulsanteReport, $pulsanteSpazio, $pulsanteCartellaReport)) { $comandiVerifiche.Controls.Add($c) }
    $script:GrigliaVerifiche = New-Griglia
    $layoutVerifiche.Controls.Add($comandiVerifiche, 0, 0)
    $layoutVerifiche.Controls.Add($script:GrigliaVerifiche, 0, 1)
    $schedaVerifiche.Controls.Add($layoutVerifiche)

    foreach ($s in @($schedaPc, $schedaRete, $schedaVerifiche)) { $schede.TabPages.Add($s) }

    $barra = New-Object System.Windows.Forms.StatusStrip
    $script:Stato = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:Stato.Text = 'Pronto.'
    [void]$barra.Items.Add($script:Stato)

    $finestra.Controls.Add($barra)
    $finestra.Controls.Add($alto)
    $finestra.Controls.Add($schede)
    $schede.BringToFront()

    $script:Pulsanti = @($pulsanteScegli, $pulsanteNuovo, $pulsanteApri, $pulsanteLeggi, $pulsanteSalva, $pulsanteScansiona, $pulsanteFileScanner, $pulsanteVerifica, $pulsanteSpazio, $pulsanteCartellaReport, $script:PulsanteAggiungi, $script:PulsanteReport)

    # Azioni
    $pulsanteScegli.Add_Click({
            try {
                $d = New-Object System.Windows.Forms.OpenFileDialog
                $d.Title = 'Excel del censimento'
                $d.Filter = 'Excel (*.xlsx)|*.xlsx'
                $d.InitialDirectory = & $script:Modulo { Get-HDCartella -Tipo Censimento }
                if ($d.ShowDialog($script:Finestra) -eq [System.Windows.Forms.DialogResult]::OK) { Update-FileCensimento $d.FileName }
            }
            catch { Show-Errore $_.Exception.Message }
        })
    $pulsanteNuovo.Add_Click({
            try {
                $cliente = [Microsoft.VisualBasic.Interaction]::InputBox('Nome del cliente:', 'Nuovo censimento', '')
                if ($cliente.Trim()) {
                    $f = New-Censimento -Cliente $cliente.Trim() -Confirm:$false
                    Update-FileCensimento $f.FullName
                    $script:Stato.Text = 'Creato ' + $f.FullName
                }
            }
            catch { Show-Errore $_.Exception.Message }
        })
    $pulsanteApri.Add_Click({
            if ($script:FileCensimento) {
                try { Invoke-Item -LiteralPath $script:FileCensimento } catch { Show-Errore $_.Exception.Message }
            }
            else { Show-Info 'Non c''è ancora un censimento: crealo con "Nuovo..."' }
        })
    $pulsanteLeggi.Add_Click({ Show-DatiPc })
    $pulsanteSalva.Add_Click({ Save-DatiPc })
    $pulsanteScansiona.Add_Click({ Start-Scansione })
    $script:PulsanteAggiungi.Add_Click({ Add-ScansioneAlCensimento })
    $pulsanteFileScanner.Add_Click({ Import-FileScanner })
    $pulsanteVerifica.Add_Click({ Start-Verifica })
    $script:PulsanteReport.Add_Click({ if ($script:UltimoReport) { Invoke-Item -LiteralPath $script:UltimoReport } })
    $pulsanteSpazio.Add_Click({ Show-SpazioDisco })
    $pulsanteCartellaReport.Add_Click({ try { Invoke-Item -LiteralPath (& $script:Modulo { Get-HDCartella -Tipo Report }) } catch { Show-Errore $_.Exception.Message } })

    $finestra.Add_Shown({
            try { Update-FileCensimento (& $script:Modulo { Resolve-HDFileCensimento -NonChiedere }) } catch { Update-FileCensimento '' }
            try {
                $locale = & $script:Modulo { Get-HDReteLocale }
                $script:CasellaRete.Text = $locale.Rete
                if ($locale.Nota) { $script:EtichettaScan.Text = $locale.Nota }
            }
            catch { $script:CasellaRete.Text = '' }
            Show-DatiPc
        })

    if ($Verifica) {
        $controlli = @($script:GrigliaPc, $script:GrigliaRete, $script:GrigliaVerifiche, $script:CasellaFile, $script:Stato)
        if (@($controlli | Where-Object { $null -eq $_ }).Count -gt 0) { throw 'Finestra incompleta' }
        $finestra.Dispose()
        exit 0
    }
    [void]$finestra.ShowDialog()
    $finestra.Dispose()
}
catch {
    Show-Errore ('La finestra non si è aperta: {0}' -f $_.Exception.Message)
    exit 1
}

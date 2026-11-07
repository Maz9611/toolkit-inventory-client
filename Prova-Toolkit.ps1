#Requires -Version 5.1
<#
.SYNOPSIS
Prova completa del toolkit da eseguire su una MACCHINA DI PROVA prima di usarlo dai clienti.

.DESCRIPTION
Esegue ogni strumento sul PC reale e controlla che funzioni senza errori.
Non modifica il PC e non tocca i censimenti veri: tutto quello che scrive finisce in
%TEMP%\HelpDeskToolkit-Prova\<data-ora> (anche un Excel di censimento di prova).
Alla fine salva un log da conservare (da allegare quando segnali un problema).

Eseguilo due volte: da utente normale e "come amministratore".

.EXAMPLE
.\Prova-Toolkit.ps1

.EXAMPLE
.\Prova-Toolkit.ps1 -SenzaInternet -SaltaPester
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Invoke-Pester -Output richiede Pester 5, caricato esplicitamente prima della chiamata')]
[CmdletBinding()]
param(
    [switch]$SenzaInternet,
    [switch]$SaltaPester,
    [switch]$NonChiedereAmministratore
)

$ErrorActionPreference = 'Stop'

# La prova completa va fatta come amministratore (SMART, TPM e BitLocker): lo propone subito
$sonoAmministratore = $false
try { $sonoAmministratore = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
catch { $sonoAmministratore = $false }
if (-not $sonoAmministratore -and -not $NonChiedereAmministratore -and [Environment]::UserInteractive) {
    Write-Host ''
    Write-Host ' Eseguire la prova come amministratore?' -ForegroundColor Cyan
    Write-Host ' Facendo questo permetterai la lettura dello stato SMART dei dischi, TPM e BitLocker.'
    Write-Host ' In caso negativo funzionerà comunque, ma senza questi dati.'
    $scelta = Read-Host ' Eseguire come amministratore? [S/n]'
    if ($scelta -notmatch '^[nN]') {
        try {
            $argomenti = @('-NoProfile', '-NoExit', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
            if ($SenzaInternet) { $argomenti += '-SenzaInternet' }
            if ($SaltaPester) { $argomenti += '-SaltaPester' }
            Start-Process -FilePath 'powershell.exe' -ArgumentList $argomenti -Verb RunAs -ErrorAction Stop
            Write-Host ' La prova continua nella nuova finestra (amministratore): questa si può chiudere.' -ForegroundColor Green
            return
        }
        catch { Write-Host ' Avvio come amministratore annullato: continuo da utente normale.' -ForegroundColor Yellow }
    }
}
$radice = $PSScriptRoot
$cartellaProva = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'HelpDeskToolkit-Prova') (Get-Date -Format 'yyyyMMdd-HHmmss')
[void][System.IO.Directory]::CreateDirectory($cartellaProva)

# Configurazione di prova: tutto quello che il toolkit scrive va nella cartella temporanea
$configProva = Join-Path $cartellaProva 'config-prova.psd1'
$q = { param($s) $s.Replace("'", "''") }
$testoConfig = "@{{`r`n    CartellaReport = '{0}'`r`n    Censimento = @{{ Cartella = '{1}' }}`r`n}}`r`n" -f (& $q (Join-Path $cartellaProva 'Report')), (& $q (Join-Path $cartellaProva 'Censimento'))
[System.IO.File]::WriteAllText($configProva, $testoConfig, (New-Object System.Text.UTF8Encoding($true)))
$configPrecedente = $env:HELPDESK_TOOLKIT_CONFIG
$env:HELPDESK_TOOLKIT_CONFIG = $configProva

$risultati = New-Object System.Collections.Generic.List[object]

function Invoke-Collaudo {
    param([Parameter(Mandatory)][string]$Nome, [Parameter(Mandatory)][scriptblock]$Blocco)
    Write-Host ''
    Write-Host "=== $Nome ===" -ForegroundColor Cyan
    $cronometro = [System.Diagnostics.Stopwatch]::StartNew()
    $esito = 'OK'
    $dettaglio = ''
    try {
        $dettaglio = (& $Blocco | Out-String).Trim()
        if ($dettaglio -like 'SALTATA*') { $esito = 'SALTATA' }
    }
    catch {
        $esito = 'FALLITA'
        $dettaglio = '{0} {1}' -f $_.Exception.Message, $_.InvocationInfo.PositionMessage
    }
    $cronometro.Stop()
    $risultati.Add([pscustomobject]@{
            Prova     = $Nome
            Esito     = $esito
            Secondi   = [math]::Round($cronometro.Elapsed.TotalSeconds, 1)
            Dettaglio = $dettaglio
        })
    $colore = 'Green'
    if ($esito -eq 'FALLITA') { $colore = 'Red' } elseif ($esito -eq 'SALTATA') { $colore = 'Yellow' }
    Write-Host ('  -> {0} ({1} s)' -f $esito, [math]::Round($cronometro.Elapsed.TotalSeconds, 1)) -ForegroundColor $colore
    if ($dettaglio) { Write-Host ('     ' + ($dettaglio -replace "`n", "`n     ")) -ForegroundColor DarkGray }
}

function Assert-Vero {
    param([bool]$Condizione, [string]$Messaggio)
    if (-not $Condizione) { throw $Messaggio }
}

Write-Host 'PROVA DEL TOOLKIT HELP DESK' -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "File di prova in: $cartellaProva"

Invoke-Collaudo 'Ambiente' {
    $os = (Get-CimInstance Win32_OperatingSystem).Caption
    $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $modo = $ExecutionContext.SessionState.LanguageMode
    'Windows: {0}' -f $os
    'PowerShell: {0} {1}' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition
    'Amministratore: {0}' -f $admin
    'Lingua/cultura: {0}' -f (Get-Culture).Name
    'Modalità linguaggio: {0}' -f $modo
    if ([string]$modo -ne 'FullLanguage') { throw "PowerShell è in modalità '$modo' (AppLocker/WDAC): il toolkit potrebbe non funzionare." }
}

Invoke-Collaudo 'Caricamento del modulo' {
    Import-Module (Join-Path (Join-Path $radice 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
    $comandi = @(Get-Command -Module HelpDeskToolkit)
    Assert-Vero ($comandi.Count -ge 9) "Attesi almeno 9 comandi, trovati $($comandi.Count)"
    ($comandi | ForEach-Object { $_.Name } | Sort-Object) -join ', '
}

$script:pc = $null
Invoke-Collaudo 'Dati di questo PC per il censimento' {
    $r = Get-InventarioPC
    Assert-Vero ($null -ne $r) 'Nessun dato raccolto'
    Assert-Vero ($r.Hostname -eq $env:COMPUTERNAME) "Hostname inatteso: $($r.Hostname)"
    Assert-Vero ([bool]$r.OS) 'Sistema operativo non letto'
    Assert-Vero ([bool]$r.Memory) 'Memoria non letta'
    Assert-Vero ([bool]$r.Processor) 'Processore non letto'
    $script:pc = $r
    $colonne = 'Hostname', 'IndirizzoIP', 'TipoIP', 'MAC', 'DeviceType', 'Manufacturer', 'Model', 'NomeUtente', 'Dominio', 'SerialNumber', 'OS', 'OSVersion', 'Memory', 'Processor', 'FreeDiskPct', 'DiskSize', 'Cynet', 'BitLocker'
    foreach ($c in $colonne) { '{0,-14} {1}' -f $c, $r.$c }
    'Nome utente da: {0}' -f $r.FonteNomeUtente
    'Cynet: {0}' -f $r.CynetDettaglio
    $vuoti = @($colonne | Where-Object { [string]::IsNullOrWhiteSpace([string]$r.$_) })
    if ($vuoti.Count -gt 0) { 'Campi vuoti (controlla se è normale su questo PC): ' + ($vuoti -join ', ') }
    Assert-Vero ($null -ne $r.Dettagli) 'Dati del foglio Dettagli non raccolti'
    ''
    'Foglio Dettagli:'
    foreach ($k in 'Windows11', 'TPM', 'SecureBoot', 'Avvio', 'DischiFisici', 'Smart', 'Batteria', 'Bios', 'Monitor', 'Stampanti', 'Office', 'AttivazioneWindows', 'Aggiornamento', 'Antivirus', 'Amministratori', 'AccessoRemoto', 'Programmi') { '{0,-19}{1}' -f $k, $r.Dettagli[$k] }
    $vuotiDettagli = @('TPM', 'Avvio', 'DischiFisici', 'Batteria', 'Bios', 'AttivazioneWindows', 'Aggiornamento', 'Amministratori', 'Programmi' | Where-Object { -not $r.Dettagli[$_] })
    if ($vuotiDettagli.Count -gt 0) { 'Dettagli vuoti (controlla se è normale su questo PC): ' + ($vuotiDettagli -join ', ') }
    'Programmi trovati: {0} (primi 5: {1})' -f @($r.Software).Count, ((@($r.Software) | Select-Object -First 5 | ForEach-Object { $_.Nome }) -join '; ')
    if (-not $r.Dettagli['Smart']) { 'SMART non letto: è normale da utente normale, rifai la prova come amministratore.' }
    if ($r.TempiRaccolta) {
        $t = $r.TempiRaccolta
        'Tempo di lettura: {0:N1} s in tutto (in parallelo: licenza {1:N1} s, programmi {2:N1} s, dischi {3:N1} s, aggiornamenti {4:N1} s, amministratori {5:N1} s; il resto {6:N1} s)' -f ($t.Totale / 1000), ($t.Licenza / 1000), ($t.Programmi / 1000), ($t.Dischi / 1000), ($t.Aggiornamenti / 1000), ($t.Amministratori / 1000), ($t.Principale / 1000)
    }
}

$script:excelProva = $null
Invoke-Collaudo 'Censimento su un Excel di prova (scansione + PC)' {
    Assert-Vero ($null -ne $script:pc) 'Manca la prova precedente (dati del PC)'
    $ip = $script:pc.IndirizzoIP
    if (-not $ip) { $ip = '192.0.2.10' }
    $scansione = Join-Path $cartellaProva 'scansione-prova.csv'
    [System.IO.File]::WriteAllLines($scansione, [string[]]@(
            'Status,Name,IP,Manufacturer,MAC address,Comments'
            ('On,{0},{1},Scheda di rete,{2},' -f $script:pc.Hostname, $ip, $script:pc.MAC)
            'On,STAMPANTE-PROVA,192.0.2.50,Stampanti di prova,02:00:00:00:00:50,'
        ))
    $import = Import-ScansioneRete -Percorso $scansione -Cliente 'Prova' -NonChiedere
    Assert-Vero ($import.Aggiunti -eq 2) "Dalla scansione attesi 2 dispositivi, aggiunti $($import.Aggiunti)"
    $esito = Add-PCCensimento -Censimento $import.Censimento -NonChiedere -Utilizzatore 'Utente di prova' -Locazione 'Banco prova'
    Assert-Vero $esito.Scritto "Excel non aggiornato: $($esito.Errore)"
    Assert-Vero ($esito.Esiti[0].Azione -eq 'Aggiornato') 'Il PC non è stato ritrovato nella riga della scansione: si è creato un doppione'
    $tabella = & (Get-Module HelpDeskToolkit) { param($p) Get-HDTabellaCensimento $p } $import.Censimento
    Assert-Vero ($tabella.Righe.Count -eq 2) "Righe attese 2, trovate $($tabella.Righe.Count)"
    $riga = $tabella.Righe | Where-Object { $_.Valori['Hostname'] -eq $script:pc.Hostname } | Select-Object -First 1
    Assert-Vero ($riga.Valori['Locazione'] -eq 'Banco prova') 'Locazione non scritta'
    Assert-Vero ($tabella.Dettagli.Esiste -and $tabella.Dettagli.Righe.Count -ge 1) 'Foglio Dettagli non scritto'
    Assert-Vero ($tabella.Riepilogo.Esiste -and -not $tabella.Riepilogo.Conflitto) 'Foglio Riepilogo non scritto'
    Assert-Vero ($tabella.Storico.Esiste -and -not $tabella.Storico.Conflitto) 'Foglio Storico mancante'
    $righeRiepilogo = @(& (Get-Module HelpDeskToolkit) { param($t) Get-HDRigheRiepilogoTabella -Tabella $t -Soglie (Get-HDConfig).Soglie } $tabella)
    $rigaPc = @($righeRiepilogo | Where-Object { $_[0] -eq $script:pc.Hostname }) | Select-Object -First 1
    if ($rigaPc) { 'Riepilogo di questo PC - urgente: {0} | da controllare: {1}' -f $(if ($rigaPc[2]) { $rigaPc[2] } else { 'niente' }), $(if ($rigaPc[3]) { $rigaPc[3] } else { 'niente' }) }
    $pagina = Export-RiepilogoCensimento -Censimento $import.Censimento -Confirm:$false
    Assert-Vero ($pagina -and (Test-Path -LiteralPath $pagina.FullName)) 'Riepilogo per il cliente non creato'
    'Riepilogo per il cliente: {0}' -f $pagina.FullName
    Assert-Vero ($riga.Valori['SerialNumber'] -eq $script:pc.SerialNumber -or -not $script:pc.SerialNumber) 'Numero di serie non scritto'
    $script:excelProva = $import.Censimento
    'Excel di prova: {0}' -f $import.Censimento
    'Riga {0}: {1}' -f $riga.Numero, (($riga.Valori.GetEnumerator() | Sort-Object Key | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ')
}

Invoke-Collaudo 'Scansione della rete (solo lettura, niente Excel)' {
    $s = Invoke-ScansioneRete -Confirm:$false
    Assert-Vero ($s.Dispositivi.Count -gt 0) "Nessun dispositivo trovato sulla rete $($s.Rete)"
    'Rete {0}: {1} dispositivi in {2} secondi' -f $s.Rete, $s.Dispositivi.Count, $s.Secondi
    $conNome = @($s.Dispositivi | Where-Object { $_.Hostname }).Count
    $conMac = @($s.Dispositivi | Where-Object { $_.MAC }).Count
    'Con nome: {0}   con MAC: {1}' -f $conNome, $conMac
    ($s.Dispositivi | Select-Object -First 15 | ForEach-Object { '{0,-15} {1,-18} {2,-17} {3}' -f $_.IP, $_.Hostname, $_.MAC, $_.Fonte }) -join "`n"
}

Invoke-Collaudo 'Finestra del toolkit (costruita e chiusa senza mostrarla)' {
    $script = Join-Path $radice 'Interfaccia.ps1'
    $uscita = & powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File $script -Verifica 2>&1
    Assert-Vero ($LASTEXITCODE -eq 0) ("La finestra non si costruisce (codice {0}): {1}" -f $LASTEXITCODE, ($uscita | Out-String).Trim())
    'La finestra si costruisce correttamente'
}

Invoke-Collaudo 'Programma ToolkitHelpDesk.exe (avviato e chiuso senza mostrarlo)' {
    $exe = Join-Path $radice 'ToolkitHelpDesk.exe'
    if (-not (Test-Path -LiteralPath $exe)) { return 'SALTATA: in questa cartella non c''è ToolkitHelpDesk.exe (si crea con Crea-Exe.cmd)' }
    $versione = (Get-Item -LiteralPath $exe).VersionInfo.FileVersion
    # avviato direttamente (non tramite la shell), così l'avviso di sicurezza per i file scaricati non interviene
    $info = New-Object System.Diagnostics.ProcessStartInfo($exe, '/verifica')
    $info.UseShellExecute = $false
    $info.WorkingDirectory = $radice
    $processo = [System.Diagnostics.Process]::Start($info)
    if (-not $processo.WaitForExit(120000)) { $processo.Kill(); throw 'L''exe non ha risposto entro 2 minuti' }
    Assert-Vero ($processo.ExitCode -eq 0) ("L'exe non si avvia (codice {0}). Dettagli in {1}" -f $processo.ExitCode, (Join-Path ([System.IO.Path]::GetTempPath()) 'HelpDeskToolkit-errori.txt'))
    'ToolkitHelpDesk.exe {0}: carica il toolkit e costruisce la finestra' -f $versione
}

Invoke-Collaudo 'Spazio disco' {
    $v = @(Get-SpazioDisco)
    Assert-Vero ($v.Count -gt 0) 'Nessun volume trovato'
    foreach ($x in $v) { Assert-Vero ($x.TotaleGB -gt 0 -and $x.Stato -in 'OK', 'ATTENZIONE', 'CRITICO') "Dati strani per $($x.Unita)" }
    ($v | ForEach-Object { '{0} {1} GB liberi su {2} GB -> {3}' -f $_.Unita, $_.LiberoGB, $_.TotaleGB, $_.Stato }) -join "`n"
}

Invoke-Collaudo 'Stima spazio recuperabile' {
    $s = @(Get-StimaPulizia)
    ($s | ForEach-Object { '{0}: {1} {2}' -f $_.Voce, $_.Dimensione, $_.Nota }) -join "`n"
}

Invoke-Collaudo 'Cartelle più grandi (profilo utente)' {
    $c = @(Get-CartelleGrandi -Percorso $env:USERPROFILE -Primi 3)
    Assert-Vero ($c.Count -gt 0) 'Nessuna cartella misurata'
    ($c | ForEach-Object { '{0} {1}' -f $_.Cartella, $_.Dimensione }) -join "`n"
}

Invoke-Collaudo 'Verifica postazione' {
    $e = @(Test-Postazione -PassThru -SenzaInternet:$SenzaInternet)
    Assert-Vero ($e.Count -ge 16) "Attesi almeno 16 controlli, ottenuti $($e.Count)"
    $errori = @($e | Where-Object { $_.Stato -eq 'ERRORE' })
    Assert-Vero ($errori.Count -eq 0) ('Controlli andati in errore: ' + (($errori | ForEach-Object { "$($_.Controllo): $($_.Dettaglio)" }) -join ' | '))
    $html = @(Get-ChildItem -LiteralPath (Join-Path $cartellaProva 'Report') -Filter 'Verifica_*.html')
    Assert-Vero ($html.Count -gt 0) 'Report HTML non creato'
    ($e | Group-Object Stato | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name }) -join ', '
    $nv = @($e | Where-Object { $_.Stato -eq 'NON VERIFICATO' })
    if ($nv.Count -gt 0) { 'Non verificati: ' + (($nv | ForEach-Object { "$($_.Controllo) ($($_.Dettaglio))" }) -join '; ') }
}

Invoke-Collaudo 'Test automatici (Pester)' {
    if ($SaltaPester) { return 'SALTATA su richiesta' }
    $pester = Get-Module -ListAvailable Pester | Where-Object { $_.Version.Major -ge 5 } | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $pester) { return 'SALTATA: Pester 5 non installato (Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck)' }
    Import-Module $pester -Force
    $r = Invoke-Pester -Path (Join-Path $radice 'tests') -PassThru -Output None
    Assert-Vero ($r.FailedCount -eq 0) ("{0} test falliti: {1}" -f $r.FailedCount, (($r.Failed | ForEach-Object { $_.ExpandedPath }) -join '; '))
    '{0} test superati' -f $r.PassedCount
}

$env:HELPDESK_TOOLKIT_CONFIG = $configPrecedente

# Riepilogo e log
$falliti = @($risultati | Where-Object { $_.Esito -eq 'FALLITA' })
Write-Host ''
Write-Host 'RIEPILOGO' -ForegroundColor White -BackgroundColor DarkBlue
foreach ($r in $risultati) {
    $colore = 'Green'
    if ($r.Esito -eq 'FALLITA') { $colore = 'Red' } elseif ($r.Esito -eq 'SALTATA') { $colore = 'Yellow' }
    Write-Host ('  {0,-8} {1,6} s  {2}' -f $r.Esito, $r.Secondi, $r.Prova) -ForegroundColor $colore
}
$log = Join-Path $cartellaProva ('Prova_{0}_{1}.txt' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmm'))
$testoLog = ($risultati | Format-List Prova, Esito, Secondi, Dettaglio | Out-String -Width 300)
[System.IO.File]::WriteAllText($log, $testoLog, (New-Object System.Text.UTF8Encoding($true)))
Write-Host ''
Write-Host "Log completo: $log"
if ($falliti.Count -eq 0) {
    Write-Host 'Tutte le prove sono andate a buon fine su questa macchina.' -ForegroundColor Green
}
else {
    Write-Host "$($falliti.Count) prove fallite: NON usare il toolkit dai clienti finché non sono risolte." -ForegroundColor Red
    Write-Host 'Suggerimento: apri il log e allega il dettaglio dell''errore quando segnali il problema.' -ForegroundColor Yellow
}
if ($script:excelProva) {
    Write-Host ''
    Write-Host 'ULTIMO CONTROLLO A MANO: apri l''Excel di prova. Excel deve aprirlo senza messaggi del tipo' -ForegroundColor Yellow
    Write-Host '"Abbiamo riscontrato un problema in alcuni contenuti". Controlla anche che i dati siano nelle colonne giuste.' -ForegroundColor Yellow
    $r = Read-Host 'Aprire adesso l''Excel di prova? [S/n]'
    if ($r -notmatch '^(n|no)$') { Invoke-Item -LiteralPath $script:excelProva }
}
if ($falliti.Count -gt 0) { exit 1 }

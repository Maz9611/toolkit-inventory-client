# Verifica rapida della postazione: comando pubblico, stampa a video e report HTML/JSON.

function Get-HDIntestazioneVerifica {
    param([Parameter(Mandatory)]$Dati)
    $modello = ''
    $dominio = ''
    $serie = ''
    if ($Dati.CS) {
        $modello = Join-HDProduttoreModello $Dati.CS.Manufacturer $Dati.CS.Model
        if ($Dati.CS.PartOfDomain) { $dominio = [string]$Dati.CS.Domain } else { $dominio = 'Nessuno (gruppo di lavoro)' }
    }
    if ($Dati.Bios) { $serie = Get-HDNumeroSerie $Dati.Bios.SerialNumber }
    $ip = @(@($Dati.Rete) | Where-Object { $_ } | ForEach-Object { @($_.IPAddress) -match $script:HDIpv4 } | Where-Object { $_ -notlike '169.254.*' }) -join ', '
    $admin = 'No'
    if ($Dati.Admin) { $admin = 'Sì' }
    [pscustomobject]@{
        NomePC         = $Dati.NomePC
        Utente         = $Dati.Utente
        Dominio        = $dominio
        Modello        = $modello
        NumeroSerie    = $serie
        IndirizzoIP    = $ip
        Amministratore = $admin
        Data           = $Dati.Raccolto
        PowerShell     = $PSVersionTable.PSVersion.ToString()
    }
}

function Get-HDConteggioStati {
    param($Esiti)
    $ordine = 'CRITICO', 'ATTENZIONE', 'ERRORE', 'NON VERIFICATO', 'INFO', 'OK'
    $conteggio = [ordered]@{}
    foreach ($s in $ordine) {
        $n = @($Esiti | Where-Object { $_.Stato -eq $s }).Count
        if ($n -gt 0) { $conteggio[$s] = $n }
    }
    $conteggio
}

function Show-HDVerifica {
    param($Intestazione, $Esiti, [string]$FileReport)
    $colori = @{ 'OK' = 'Green'; 'ATTENZIONE' = 'Yellow'; 'CRITICO' = 'Red'; 'INFO' = 'Cyan'; 'NON VERIFICATO' = 'DarkGray'; 'ERRORE' = 'Magenta' }
    Write-Host ''
    Write-Host (' VERIFICA POSTAZIONE  {0}  -  {1} ' -f $Intestazione.NomePC, (Format-HDData $Intestazione.Data)) -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host (' Utente: {0}   Dominio: {1}' -f $Intestazione.Utente, $Intestazione.Dominio)
    Write-Host (' Modello: {0}   S/N: {1}   IP: {2}' -f $Intestazione.Modello, $Intestazione.NumeroSerie, $Intestazione.IndirizzoIP)
    if ($Intestazione.Amministratore -ne 'Sì') {
        Write-Host ' Non sei amministratore: alcuni controlli (BitLocker, punti di ripristino) sono limitati.' -ForegroundColor DarkGray
    }
    Write-Host (' ' + ('-' * 76))
    foreach ($e in $Esiti) {
        Write-Host (' {0,-17}' -f ('[' + $e.Stato + ']')) -ForegroundColor $colori[$e.Stato] -NoNewline
        Write-Host ('{0,-24}' -f $e.Controllo) -NoNewline
        Write-Host $e.Dettaglio
        if ($e.Suggerimento -and $e.Stato -ne 'OK') {
            Write-Host ((' ' * 18) + '-> ' + $e.Suggerimento) -ForegroundColor DarkYellow
        }
    }
    Write-Host (' ' + ('-' * 76))
    $conteggio = Get-HDConteggioStati $Esiti
    Write-Host ' Riepilogo: ' -NoNewline
    foreach ($k in $conteggio.Keys) { Write-Host ('{0} {1}   ' -f $conteggio[$k], $k) -ForegroundColor $colori[$k] -NoNewline }
    Write-Host ''
    if ($FileReport) { Write-Host " Report salvato: $FileReport" -ForegroundColor Green }
    Write-Host ''
}

function ConvertTo-HDHtmlVerifica {
    param([Parameter(Mandatory)]$Intestazione, [Parameter(Mandatory)]$Esiti, [string]$Azienda)
    $campi = [ordered]@{
        'PC'             = $Intestazione.NomePC
        'Utente'         = $Intestazione.Utente
        'Modello'        = $Intestazione.Modello
        'N. di serie'    = $Intestazione.NumeroSerie
        'Dominio'        = $Intestazione.Dominio
        'Indirizzo IP'   = $Intestazione.IndirizzoIP
        'Amministratore' = $Intestazione.Amministratore
        'PowerShell'     = $Intestazione.PowerShell
    }
    $griglia = ($campi.Keys | ForEach-Object {
            '<div class="campo"><span class="etichetta">{0}</span><span>{1}</span></div>' -f (ConvertTo-HDHtml $_), (ConvertTo-HDHtml ([string]$campi[$_]))
        }) -join "`n"
    $conteggio = Get-HDConteggioStati $Esiti
    $riepilogo = ($conteggio.Keys | ForEach-Object { '<span class="badge {0}">{1} {2}</span>' -f (Get-HDClasseStato $_), $conteggio[$_], (ConvertTo-HDHtml $_) }) -join ' '
    $righe = foreach ($e in $Esiti) {
        $sugg = ''
        if ($e.Suggerimento -and $e.Stato -ne 'OK') { $sugg = '<div class="tenue">&rarr; {0}</div>' -f (ConvertTo-HDHtml $e.Suggerimento) }
        '<tr><td><strong>{0}</strong><div class="tenue">{1}</div></td><td><span class="badge {2}">{3}</span></td><td>{4}{5}</td></tr>' -f `
        (ConvertTo-HDHtml $e.Controllo), (ConvertTo-HDHtml $e.Area), (Get-HDClasseStato $e.Stato), (ConvertTo-HDHtml $e.Stato), (ConvertTo-HDHtml $e.Dettaglio), $sugg
    }
    $data = Format-HDData $Intestazione.Data
    $corpo = @"
<div class="azioni no-print"><button type="button" onclick="window.print()">Stampa / Salva PDF</button></div>
<div class="intestazione">
  <div><h1>Verifica postazione</h1><div class="tenue">$(ConvertTo-HDHtml $Azienda)</div></div>
  <div class="destra"><strong>$(ConvertTo-HDHtml $Intestazione.NomePC)</strong><br>$data</div>
</div>
<div class="griglia">
$griglia
</div>
<h2>Riepilogo</h2>
<p>$riepilogo</p>
<h2>Controlli</h2>
<table>
<thead><tr><th>Controllo</th><th>Stato</th><th>Dettaglio</th></tr></thead>
<tbody>
$($righe -join "`n")
</tbody>
</table>
<div class="piede">Generato il $data con HelpDesk Toolkit &middot; verifica in sola lettura, nessuna modifica al PC.</div>
"@
    ConvertTo-HDPaginaHtml -Titolo ('Verifica {0}' -f $Intestazione.NomePC) -Corpo $corpo
}

function Test-Postazione {
    <#
    .SYNOPSIS
    Verifica rapida della postazione da fare PRIMA di un intervento.

    .DESCRIPTION
    Controlla in pochi secondi: Windows e supporto, ultimo riavvio, riavvii in sospeso,
    spazio disco, RAM, aggiornamenti, antivirus, firewall, BitLocker, rete/DNS/Internet,
    orologio, errori nel Registro eventi, salute dei dischi, servizi essenziali,
    stampanti e punti di ripristino. Mostra il risultato a colori e salva un report
    HTML (stampabile) e JSON nella cartella dei report. Non modifica nulla.

    .PARAMETER SenzaInternet
    Salta le prove verso Internet (DNS pubblico, sito di prova, confronto orario).

    .PARAMETER Escludi
    Controlli da saltare, ad esempio -Escludi Eventi, Stampanti.

    .PARAMETER NoReport
    Non salva i file HTML/JSON.

    .PARAMETER PassThru
    Restituisce anche gli esiti come oggetti (per filtrarli o esportarli).

    .PARAMETER Silenzioso
    Non stampa a video.

    .EXAMPLE
    Test-Postazione

    .EXAMPLE
    Test-Postazione -PassThru -Silenzioso | Where-Object Stato -in 'CRITICO', 'ATTENZIONE'
    #>
    [CmdletBinding()]
    param(
        [switch]$SenzaInternet,
        [ValidateSet('SistemaOperativo', 'Uptime', 'Riavvio', 'Disco', 'Memoria', 'Aggiornamenti', 'Antivirus', 'Firewall',
            'BitLocker', 'Rete', 'Ora', 'Eventi', 'DischiFisici', 'Servizi', 'Stampanti', 'PuntoRipristino')]
        [string[]]$Escludi,
        [switch]$NoReport,
        [string]$CartellaOutput,
        [switch]$PassThru,
        [switch]$Silenzioso
    )
    if (-not (Test-HDWindows)) { throw 'La verifica della postazione funziona solo su Windows.' }
    $cfg = Get-HDConfig
    $dati = Get-HDDatiPostazione -SenzaInternet:$SenzaInternet
    $esiti = @(Invoke-HDControlli -Dati $dati -Soglie $cfg.Soglie -Adesso $dati.Raccolto -Escludi $Escludi)
    $intestazione = Get-HDIntestazioneVerifica -Dati $dati

    $fileHtml = $null
    if (-not $NoReport) {
        $cartella = Get-HDCartella -Tipo Report -Specificata $CartellaOutput
        $base = 'Verifica_{0}_{1}' -f (ConvertTo-HDNomeFile $intestazione.NomePC), $dati.Raccolto.ToString('yyyyMMdd-HHmm')
        $fileHtml = Get-HDPercorsoUnico -Cartella $cartella -NomeBase $base -Estensione '.html'
        $fileJson = [System.IO.Path]::ChangeExtension($fileHtml, '.json')
        Write-HDFileTesto -Percorso $fileHtml -Testo (ConvertTo-HDHtmlVerifica -Intestazione $intestazione -Esiti $esiti -Azienda $cfg.Azienda)
        $json = [pscustomobject]@{ Postazione = (ConvertTo-HDRigaTesto -Oggetto $intestazione); Controlli = $esiti } | ConvertTo-Json -Depth 5
        Write-HDFileTesto -Percorso $fileJson -Testo $json
        $script:HDUltimaVerifica = $fileHtml
    }
    if (-not $Silenzioso) { Show-HDVerifica -Intestazione $intestazione -Esiti $esiti -FileReport $fileHtml }
    if ($PassThru) { $esiti }
}

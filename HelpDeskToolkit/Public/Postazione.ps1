# Verifica rapida della postazione: comando pubblico, stampa a video e report HTML/JSON.

function Get-HDIntestazioneVerifica {
    param([Parameter(Mandatory)]$Dati)
    $modello = ''
    $dominio = ''
    $serie = ''
    if ($Dati.CS) {
        $scheda = ''
        $produttore = ConvertTo-HDProduttore ([string]$Dati.CS.Manufacturer)
        if ($Dati.SchedaMadre) {
            $scheda = [string]$Dati.SchedaMadre.Product
            if (-not $produttore) { $produttore = ConvertTo-HDProduttore ([string]$Dati.SchedaMadre.Manufacturer) }
        }
        $modello = Join-HDProduttoreModello $produttore (Get-HDModelloLeggibile -Produttore ([string]$Dati.CS.Manufacturer) -Modello ([string]$Dati.CS.Model) -SchedaMadre $scheda)
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
    # Report della verifica: esito in una frase, prima i problemi con cosa fare, poi i controlli superati
    param([Parameter(Mandatory)]$Intestazione, [Parameter(Mandatory)]$Esiti)
    $h = { param($t) ConvertTo-HDHtml ([string]$t) }
    $esiti = @($Esiti | Where-Object { $_ })
    $conteggio = Get-HDConteggioStati $esiti
    $n = { param($stato) $v = $conteggio[$stato]; if ($v) { [int]$v } else { 0 } }
    $gravi = (& $n 'CRITICO') + (& $n 'ERRORE')
    $daSeguire = & $n 'ATTENZIONE'
    if ($gravi -gt 0) {
        $frase = Format-HDQuantita $gravi 'problema da sistemare' 'problemi da sistemare'
        if ($daSeguire -gt 0) { $frase += ', ' + (Format-HDQuantita $daSeguire "punto da tenere d'occhio" "punti da tenere d'occhio") }
    }
    elseif ($daSeguire -gt 0) { $frase = 'Nessun problema grave, ' + (Format-HDQuantita $daSeguire "punto da tenere d'occhio" "punti da tenere d'occhio") }
    else { $frase = 'Nessun problema rilevato' }

    $nomi = [ordered]@{ 'CRITICO' = @('critico', 'critici'); 'ERRORE' = @('non riuscito', 'non riusciti'); 'ATTENZIONE' = @('da tenere d''occhio', 'da tenere d''occhio'); 'NON VERIFICATO' = @('non verificato', 'non verificati'); 'INFO' = @('informativo', 'informativi'); 'OK' = @('a posto', 'a posto') }
    $barra = ($nomi.Keys | Where-Object { & $n $_ } | ForEach-Object { '<span class="{0}" style="flex:{1}"></span>' -f ((Get-HDClasseStato $_) -replace '^s-', 'c-'), (& $n $_) }) -join ''
    $legenda = ($nomi.Keys | Where-Object { & $n $_ } | ForEach-Object { '<li><i class="punto {0}"></i>{1}</li>' -f ((Get-HDClasseStato $_) -replace '^s-', 'c-'), (Format-HDQuantita (& $n $_) $nomi[$_][0] $nomi[$_][1]) }) -join ''

    $campi = [ordered]@{
        'Utente'          = $Intestazione.Utente
        'Dominio'         = $Intestazione.Dominio
        'Modello'         = $Intestazione.Modello
        'Numero di serie' = $Intestazione.NumeroSerie
        'Indirizzo IP'    = $Intestazione.IndirizzoIP
        'Verifica come'   = $(if ($Intestazione.Amministratore -eq 'Sì') { 'amministratore' } else { 'utente normale' })
    }
    $dati = ($campi.Keys | Where-Object { [string]$campi[$_] } | ForEach-Object { '<div><dt>{0}</dt><dd>{1}</dd></div>' -f (& $h $_), (& $h $campi[$_]) }) -join "`n"

    $ordine = @{ 'CRITICO' = 0; 'ERRORE' = 1; 'ATTENZIONE' = 2; 'NON VERIFICATO' = 3; 'INFO' = 4; 'OK' = 5 }
    $problemi = @($esiti | Where-Object { $_.Stato -in 'CRITICO', 'ERRORE', 'ATTENZIONE' } | Sort-Object { $ordine[[string]$_.Stato] })
    $altri = @($esiti | Where-Object { $_.Stato -notin 'CRITICO', 'ERRORE', 'ATTENZIONE' } | Sort-Object { $ordine[[string]$_.Stato] })
    $voci = foreach ($e in $problemi) {
        $classe = (Get-HDClasseStato $e.Stato) -replace '^s-', ''
        $fare = ''
        if ($e.Suggerimento) { $fare = '<p class="fare"><b>Cosa fare:</b> {0}</p>' -f (& $h $e.Suggerimento) }
        '<article class="voce v-{0}"><div class="testa"><span class="nome">{1}</span><span class="stato t-{0}">{2}</span><span class="area">{3}</span></div><p>{4}</p>{5}</article>' -f `
            $classe, (& $h $e.Controllo), (& $h $e.Stato), (& $h $e.Area), ((& $h $e.Dettaglio) -replace ' \| ', ' &middot; '), $fare
    }
    $righe = foreach ($e in $altri) {
        $classe = (Get-HDClasseStato $e.Stato) -replace '^s-', ''
        $nota = ''
        if ($e.Stato -eq 'NON VERIFICATO') { $nota = ' <span class="t-{0}">(non verificato)</span>' -f $classe }
        if ($e.Suggerimento -and $e.Stato -ne 'OK') { $nota += '<div class="nota">{0}</div>' -f (& $h $e.Suggerimento) }
        '<tr><td class="controllo"><i class="punto c-{0}"></i>{1}</td><td>{2}{3}</td><td class="area">{4}</td></tr>' -f $classe, (& $h $e.Controllo), ((& $h $e.Dettaglio) -replace ' \| ', ' &middot; '), $nota, (& $h $e.Area)
    }
    $sezioneProblemi = ''
    if ($problemi.Count -gt 0) { $sezioneProblemi = "<h2>Da sistemare</h2>`n" + ($voci -join "`n") }
    $sezioneAltri = ''
    if ($altri.Count -gt 0) { $sezioneAltri = "<h2>Gli altri controlli</h2>`n<table><thead><tr><th>Controllo</th><th>Risultato</th><th>Area</th></tr></thead><tbody>`n" + ($righe -join "`n") + "`n</tbody></table>" }
    $pezzi = @($Intestazione.Modello)
    if ($Intestazione.NumeroSerie) { $pezzi += 'S/N ' + $Intestazione.NumeroSerie }
    $sottotitolo = (@($pezzi | Where-Object { $_ } | ForEach-Object { & $h $_ }) -join ' &middot; ')
    $giorno = Format-HDData $Intestazione.Data 'dd/MM/yyyy'
    $ora = Format-HDData $Intestazione.Data 'HH:mm'

    $corpo = @"
<div class="azioni no-print"><button type="button" onclick="window.print()">Stampa o salva PDF</button></div>
<main class="pagina">
<header class="testata">
  <div>
    <div class="etichetta">Verifica postazione</div>
    <h1>$(& $h $Intestazione.NomePC)</h1>
    <div class="sottotitolo">$sottotitolo</div>
  </div>
  <div class="quando"><strong>$giorno</strong>ore $ora</div>
</header>
<section class="esito">
  <div class="frase">$(& $h $frase)</div>
  <div class="barra">$barra</div>
  <ul class="legenda">$legenda</ul>
</section>
<h2>Postazione</h2>
<dl class="dati">
$dati
</dl>
$sezioneProblemi
$sezioneAltri
<footer class="piede"><span>Verifica in sola lettura: il PC non è stato modificato.</span><span>Toolkit Help Desk &middot; $giorno $ora</span></footer>
</main>
"@
    ConvertTo-HDPaginaHtml -Titolo ('Verifica {0} - {1}' -f $Intestazione.NomePC, $giorno) -Corpo $corpo
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
        Write-HDFileTesto -Percorso $fileHtml -Testo (ConvertTo-HDHtmlVerifica -Intestazione $intestazione -Esiti $esiti)
        $json = [pscustomobject]@{ Postazione = (ConvertTo-HDRigaTesto -Oggetto $intestazione); Controlli = $esiti } | ConvertTo-Json -Depth 5
        Write-HDFileTesto -Percorso $fileJson -Testo $json
        $script:HDUltimaVerifica = $fileHtml
    }
    if (-not $Silenzioso) { Show-HDVerifica -Intestazione $intestazione -Esiti $esiti -FileReport $fileHtml }
    if ($PassThru) { $esiti }
}

# Riepilogo del parco PC da consegnare al cliente: pagina HTML dai dati dell'Excel del censimento.

function ConvertTo-HDHtmlRiepilogo {
    # Pagina del riepilogo: esito in una frase, punti da sistemare per priorità con i PC e cosa fare,
    # il parco in numeri, i cambiamenti recenti e l'elenco dei PC
    param([string]$Cliente, [object[]]$PC, [object[]]$Punti, [object[]]$Cambiamenti, [string]$File, [datetime]$Adesso = (Get-Date), [int]$GiorniCambiamenti = 90)
    $h = { param($t) ConvertTo-HDHtml ([string]$t) }
    $elencoPc = @($PC | Where-Object { $_ })
    $tuttiPunti = @($Punti | Where-Object { $_ })
    $priorita = [ordered]@{
        'Urgente'        = @{ Classe = 'critico'; Sezione = 'Da fare subito' }
        'Da pianificare' = @{ Classe = 'attenzione'; Sezione = 'Da pianificare' }
        'Da sapere'      = @{ Classe = 'info'; Sezione = 'Da sapere' }
    }

    # Esito: quanti punti per priorità e quanti PC coinvolti (ogni PC conta una volta, col punto più grave)
    $numeroPunti = @{}
    foreach ($k in $priorita.Keys) { $numeroPunti[$k] = @($tuttiPunti | Where-Object { $_.Priorita -eq $k }).Count }
    if ($numeroPunti['Urgente'] -gt 0) {
        $frase = Format-HDQuantita $numeroPunti['Urgente'] 'punto urgente' 'punti urgenti'
        if ($numeroPunti['Da pianificare'] -gt 0) { $frase += ' e ' + (Format-HDQuantita $numeroPunti['Da pianificare'] 'da pianificare' 'da pianificare') }
    }
    elseif ($numeroPunti['Da pianificare'] -gt 0) { $frase = 'Nessun punto urgente, ' + (Format-HDQuantita $numeroPunti['Da pianificare'] 'da pianificare' 'da pianificare') }
    else { $frase = 'Nessun punto da sistemare' }
    $peggiore = @{}
    foreach ($k in @($priorita.Keys)) {
        foreach ($p in @($tuttiPunti | Where-Object { $_.Priorita -eq $k })) {
            foreach ($x in @($p.PC)) { if (-not $peggiore.ContainsKey($x.Nome)) { $peggiore[$x.Nome] = $k } }
        }
    }
    $conta = [ordered]@{ 'Urgente' = 0; 'Da pianificare' = 0; 'Da sapere' = 0; 'OK' = 0 }
    foreach ($x in $elencoPc) { if ($peggiore.ContainsKey($x.Nome)) { $conta[$peggiore[$x.Nome]]++ } else { $conta['OK']++ } }
    $coinvolti = $conta['Urgente'] + $conta['Da pianificare']
    $sotto = ''
    if ($elencoPc.Count -gt 1 -and $coinvolti -eq $elencoPc.Count) { $sotto = 'Riguardano tutti i PC censiti.' }
    elseif ($elencoPc.Count -gt 0 -and $coinvolti -gt 0) { $sotto = 'Riguardano {0} dei {1}.' -f $coinvolti, (Format-HDQuantita $elencoPc.Count 'PC censito' 'PC censiti') }
    $classi = @{ 'Urgente' = 'critico'; 'Da pianificare' = 'attenzione'; 'Da sapere' = 'info'; 'OK' = 'ok' }
    $nomiLegenda = @{ 'Urgente' = 'con punti urgenti'; 'Da pianificare' = 'con punti da pianificare'; 'Da sapere' = 'solo da sapere'; 'OK' = 'a posto' }
    $barra = (@($conta.Keys | Where-Object { $conta[$_] -gt 0 } | ForEach-Object { '<span class="c-{0}" style="flex:{1}"></span>' -f $classi[$_], $conta[$_] }) -join '')
    $legenda = (@($conta.Keys | Where-Object { $conta[$_] -gt 0 } | ForEach-Object { '<li><i class="punto c-{0}"></i>{1} {2}</li>' -f $classi[$_], (Format-HDQuantita $conta[$_] 'PC' 'PC'), $nomiLegenda[$_] }) -join '')

    # Punti da sistemare
    $sezioni = foreach ($k in $priorita.Keys) {
        $diQuesta = @($tuttiPunti | Where-Object { $_.Priorita -eq $k })
        if ($diQuesta.Count -eq 0) { continue }
        $classe = $priorita[$k].Classe
        $voci = foreach ($p in $diQuesta) {
            $elenco = foreach ($x in @($p.PC)) {
                $riga = '<b>{0}</b>' -f (& $h $x.Nome)
                if ($x.Utilizzatore) { $riga += ' &middot; ' + (& $h $x.Utilizzatore) }
                if ($x.Dettaglio) { $riga += ' <span class="nota">{0}</span>' -f (& $h $x.Dettaglio) }
                '<li>{0}</li>' -f $riga
            }
            '<article class="voce v-{0}"><div class="testa"><span class="nome">{1}</span><span class="stato t-{0}">{2}</span></div><ul class="pc">{3}</ul><p class="fare"><b>Cosa fare:</b> {4}</p></article>' -f `
                $classe, (& $h $p.Punto), (& $h (Format-HDQuantita @($p.PC).Count 'PC' 'PC')), (@($elenco) -join ''), (& $h $p.CosaFare)
        }
        "<h2>$($priorita[$k].Sezione)</h2>`n" + (@($voci) -join "`n")
    }

    # Il parco in numeri
    $contaSe = { param($Filtro) @($elencoPc | Where-Object $Filtro).Count }
    $numeri = [ordered]@{}
    $numeri['PC censiti'] = [string]$elencoPc.Count
    $numeri['Con Windows 11'] = [string](& $contaSe { [string]$_.Valori['OS'] -match '(?i)windows 11' })
    $w10 = & $contaSe { [string]$_.Valori['OS'] -match '(?i)windows 10' }
    if ($w10 -gt 0) {
        $numeri['Con Windows 10'] = [string]$w10
        $numeri['Windows 10 compatibili con Windows 11'] = [string](& $contaSe { [string]$_.Valori['OS'] -match '(?i)windows 10' -and [string]$_.Dettagli['Windows11'] -match '^S' })
    }
    $altri = & $contaSe { [string]$_.Valori['OS'] -and [string]$_.Valori['OS'] -notmatch '(?i)windows 1[01]' }
    if ($altri -gt 0) { $numeri['Con altri sistemi'] = [string]$altri }
    $portatili = & $contaSe { [string]$_.Dettagli['Batteria'] -and [string]$_.Dettagli['Batteria'] -ne 'Nessuna' }
    $fissi = & $contaSe { [string]$_.Dettagli['Batteria'] -eq 'Nessuna' }
    if ($portatili + $fissi -gt 0) { $numeri['Portatili / fissi'] = '{0} / {1}' -f $portatili, $fissi }
    $m365 = & $contaSe { [string]$_.Dettagli['Office'] -match '365' }
    if ($m365 -gt 0) { $numeri['Con Microsoft 365'] = [string]$m365 }
    $date = @($elencoPc | ForEach-Object { ConvertFrom-HDDataTesto $_.Dettagli['Rilevato'] } | Where-Object { $_ } | Sort-Object)
    if ($date.Count -gt 0) {
        $primo = $date[0].ToString('dd/MM/yyyy', $script:HDInvariante)
        $ultimo = $date[-1].ToString('dd/MM/yyyy', $script:HDInvariante)
        if ($primo -eq $ultimo) { $numeri['Dati raccolti il'] = $primo } else { $numeri['Dati raccolti'] = "dal $primo al $ultimo" }
    }
    $dati = (@($numeri.Keys | ForEach-Object { '<div><dt>{0}</dt><dd>{1}</dd></div>' -f (& $h $_), (& $h $numeri[$_]) }) -join "`n")

    # Cambiamenti recenti (dal foglio Storico)
    $limite = $Adesso.AddDays(-1 * $GiorniCambiamenti)
    $recenti = @($Cambiamenti | Where-Object { $_ } | ForEach-Object {
            $d = ConvertFrom-HDDataTesto $_.Data
            if ($d -and $d -ge $limite) { [pscustomobject]@{ Data = $d; Voce = $_ } }
        } | Sort-Object Data -Descending | Select-Object -First 60)
    $sezioneCambiamenti = ''
    if ($recenti.Count -gt 0) {
        $righe = foreach ($r in $recenti) {
            $v = $r.Voce
            $cosa = ''
            if ($v.Prima -and $v.Dopo) { $cosa = '{0} <span class="nota">&rarr;</span> {1}' -f (& $h $v.Prima), (& $h $v.Dopo) }
            elseif ($v.Dopo) { $cosa = & $h $v.Dopo }
            else { $cosa = & $h $v.Prima }
            '<tr><td class="data">{0}</td><td class="controllo">{1}</td><td>{2}<div class="nota">{3}</div></td></tr>' -f $r.Data.ToString('dd/MM/yyyy', $script:HDInvariante), (& $h $v.Hostname), (& $h $v.Campo), $cosa
        }
        $sezioneCambiamenti = "<h2>Cambiamenti degli ultimi $GiorniCambiamenti giorni</h2>`n<table class=`"tabella-larga`"><thead><tr><th>Data</th><th>PC</th><th>Cosa è cambiato</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
    }

    # Elenco dei PC
    $sezioneElenco = ''
    if ($elencoPc.Count -gt 0) {
        $righe = foreach ($x in @($elencoPc | Sort-Object Nome)) {
            $stato = 'OK'
            if ($peggiore.ContainsKey($x.Nome)) { $stato = $peggiore[$x.Nome] }
            $sistema = (([string]$x.Valori['OS']) -replace '^(?i)Microsoft\s+', '').Trim()
            $versione = (([string]$x.Valori['OSVersion']) -replace '\s*\(.*$', '').Trim()
            if ($versione) { $sistema = "$sistema $versione" }
            $ram = ''
            $gb = Get-HDMemoriaGB $x.Valori['Memory']
            if ($gb -gt 0) { $ram = '{0} GB' -f [math]::Round($gb, 1) }
            $modello = Join-HDProduttoreModello $x.Valori['Manufacturer'] $x.Valori['Model']
            '<tr><td class="controllo"><i class="punto c-{0}"></i>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td class="num">{5}</td></tr>' -f $classi[$stato], (& $h $x.Nome), (& $h $x.Utilizzatore), (& $h $modello), (& $h $sistema), (& $h $ram)
        }
        $sezioneElenco = "<h2>Elenco dei PC</h2>`n<table class=`"tabella-larga`"><thead><tr><th>PC</th><th>Utilizzatore</th><th>Modello</th><th>Sistema</th><th>RAM</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
    }

    $titolo = $Cliente
    if (-not $titolo) { $titolo = 'Parco PC' }
    $sottotitolo = Format-HDQuantita $elencoPc.Count 'PC censito' 'PC censiti'
    $giorno = Format-HDData $Adesso 'dd/MM/yyyy'
    $sezioniPunti = @($sezioni) -join "`n"
    $paragrafoSotto = ''
    if ($sotto) { $paragrafoSotto = '<div class="sotto">{0}</div>' -f (& $h $sotto) }
    $origine = 'Dati raccolti in sola lettura dal censimento'
    if ($File) { $origine += ' ' + $File }

    $corpo = @"
<div class="azioni no-print"><button type="button" onclick="window.print()">Stampa o salva PDF</button></div>
<main class="pagina">
<header class="testata">
  <div>
    <div class="etichetta">Riepilogo del parco PC</div>
    <h1>$(& $h $titolo)</h1>
    <div class="sottotitolo">$(& $h $sottotitolo)</div>
  </div>
  <div class="quando"><strong>$giorno</strong>situazione all'ultimo censimento</div>
</header>
<section class="esito">
  <div class="frase">$(& $h $frase)</div>
  $paragrafoSotto
  <div class="barra">$barra</div>
  <ul class="legenda">$legenda</ul>
</section>
$sezioniPunti
<h2>Il parco PC in numeri</h2>
<dl class="dati">
$dati
</dl>
$sezioneCambiamenti
$sezioneElenco
<footer class="piede"><span>$(& $h $origine).</span><span>Toolkit Help Desk &middot; $giorno</span></footer>
</main>
"@
    ConvertTo-HDPaginaHtml -Titolo ('Riepilogo {0} - {1}' -f $titolo, $giorno) -Corpo $corpo
}

function Export-RiepilogoCensimento {
    <#
    .SYNOPSIS
    Crea il riepilogo del parco PC da consegnare al cliente (pagina HTML da stampare o salvare in PDF).

    .DESCRIPTION
    Legge l'Excel del censimento e mette in una pagina i punti da sistemare divisi per priorità (con i PC
    interessati e cosa fare), il parco PC in numeri, i cambiamenti degli ultimi 90 giorni (foglio Storico)
    e l'elenco dei PC. Gli stessi punti sono anche nel foglio "Riepilogo" dell'Excel, ricalcolato a ogni
    censimento. Il file va nella cartella Report del toolkit.

    .PARAMETER Censimento
    File Excel del censimento (se non indicato: quello della cartella Censimento).

    .PARAMETER Percorso
    Dove salvare la pagina (predefinito: Report\Riepilogo_<cliente>_<data>.html).

    .PARAMETER Apri
    Apre la pagina nel browser.

    .EXAMPLE
    Export-RiepilogoCensimento -Apri
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Position = 0)][string]$Censimento, [string]$Percorso, [switch]$Apri)
    $file = Resolve-HDFileCensimento -Percorso $Censimento -NonChiedere:(-not (Test-HDInterattivo))
    $tabella = Get-HDTabellaCensimento $file
    $cfg = Get-HDConfig
    $adesso = Get-Date
    $dettagli = @()
    if ($tabella.Dettagli) { $dettagli = @($tabella.Dettagli.Righe | Where-Object { $_ }) }
    $pc = @(Get-HDPcCensiti -Righe @($tabella.Righe | Where-Object { $_ }) -Dettagli $dettagli)
    $punti = @(Get-HDPuntiRiepilogo -PC $pc -Soglie $cfg.Soglie -Adesso $adesso)
    $cambiamenti = @()
    if ($tabella.Storico -and -not $tabella.Storico.Conflitto) {
        $cambiamenti = @($tabella.Storico.Righe | ForEach-Object { [pscustomobject]@{ Data = $_.Valori['Data']; Hostname = $_.Valori['Hostname']; Campo = $_.Valori['Campo']; Prima = $_.Valori['Prima']; Dopo = $_.Valori['Dopo'] } })
    }
    $cliente = Get-HDNomeCliente $file
    $html = ConvertTo-HDHtmlRiepilogo -Cliente $cliente -PC $pc -Punti $punti -Cambiamenti $cambiamenti -File (Split-Path -Leaf $file) -Adesso $adesso
    if (-not $Percorso) {
        $nome = 'Riepilogo'
        if ($cliente) { $nome += '_' + (ConvertTo-HDNomeFile $cliente) }
        $Percorso = Join-Path (Get-HDCartella -Tipo Report) ('{0}_{1}.html' -f $nome, $adesso.ToString('yyyyMMdd-HHmm', $script:HDInvariante))
    }
    $destinazione = Resolve-HDPercorso $Percorso
    if (-not $PSCmdlet.ShouldProcess($destinazione, 'Crea il riepilogo del parco PC')) { return }
    Write-HDFileTesto -Percorso $destinazione -Testo $html
    if ($Apri) { Start-Process -FilePath $destinazione }
    Get-Item -LiteralPath $destinazione
}

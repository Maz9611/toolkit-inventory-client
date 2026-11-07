# Riepilogo del parco PC da consegnare al cliente: pagina HTML dai dati dell'Excel del censimento.

function ConvertTo-HDHtmlRiepilogo {
    # Pagina del riepilogo: esito in una frase, i problemi in sintesi (quanti PC per ciascuno), PC per PC
    # con i fatti, il parco in numeri, i cambiamenti recenti e l'elenco dei PC
    param([string]$Cliente, [object[]]$PC, [object[]]$Punti, [object[]]$Cambiamenti, [string]$File, [datetime]$Adesso = (Get-Date), [int]$GiorniCambiamenti = 90)
    $h = { param($t) ConvertTo-HDHtml ([string]$t) }
    $elencoPc = @($PC | Where-Object { $_ })
    $tuttiPunti = @($Punti | Where-Object { $_ })
    $perPc = @(Get-HDRiepilogoPerPc -PC $elencoPc -Punti $tuttiPunti)
    $classi = @{ 'Urgente' = 'critico'; 'Da controllare' = 'attenzione'; 'OK' = 'ok' }
    $stato = @{}
    foreach ($x in $perPc) {
        $stato[$x.Nome] = 'OK'
        if ($x.Urgente.Count -gt 0) { $stato[$x.Nome] = 'Urgente' }
        elseif ($x.DaControllare.Count -gt 0) { $stato[$x.Nome] = 'Da controllare' }
    }

    # Esito: quanti PC hanno problemi urgenti, quanti solo punti da controllare, quanti nessuno
    $conta = [ordered]@{ 'Urgente' = 0; 'Da controllare' = 0; 'OK' = 0 }
    foreach ($k in $stato.Keys) { $conta[$stato[$k]]++ }
    if ($conta['Urgente'] -gt 0) {
        $frase = Format-HDQuantita $conta['Urgente'] 'PC con problemi urgenti' 'PC con problemi urgenti'
        if ($conta['Da controllare'] -gt 0) { $frase += ', {0} con punti da controllare' -f $conta['Da controllare'] }
    }
    elseif ($conta['Da controllare'] -gt 0) { $frase = 'Nessun problema urgente, ' + (Format-HDQuantita $conta['Da controllare'] 'PC con punti da controllare' 'PC con punti da controllare') }
    else { $frase = 'Nessun problema rilevato' }
    $sotto = ''
    if ($conta['OK'] -gt 0 -and $conta['OK'] -lt $perPc.Count) { $sotto = '{0} su {1} senza problemi.' -f (Format-HDQuantita $conta['OK'] 'PC' 'PC'), $perPc.Count }
    $nomiLegenda = @{ 'Urgente' = 'con problemi urgenti'; 'Da controllare' = 'con punti da controllare'; 'OK' = 'senza problemi' }
    $barra = (@($conta.Keys | Where-Object { $conta[$_] -gt 0 } | ForEach-Object { '<span class="c-{0}" style="flex:{1}"></span>' -f $classi[$_], $conta[$_] }) -join '')
    $legenda = (@($conta.Keys | Where-Object { $conta[$_] -gt 0 } | ForEach-Object { '<li><i class="punto c-{0}"></i>{1} {2}</li>' -f $classi[$_], (Format-HDQuantita $conta[$_] 'PC' 'PC'), $nomiLegenda[$_] }) -join '')

    # In sintesi: un problema per riga con quanti PC lo hanno, prima gli urgenti
    $sezioneSintesi = ''
    if ($tuttiPunti.Count -gt 0) {
        $ordinati = @($tuttiPunti | Sort-Object @{ Expression = { if ($_.Priorita -eq 'Urgente') { 0 } else { 1 } } }, @{ Expression = { @($_.PC).Count }; Descending = $true })
        $righe = foreach ($p in $ordinati) {
            '<tr><td class="controllo"><i class="punto c-{0}"></i>{1}</td><td class="num">{2}</td></tr>' -f $classi[$p.Priorita], (& $h $p.Punto), (& $h (Format-HDQuantita @($p.PC).Count 'PC' 'PC'))
        }
        $sezioneSintesi = "<h2>In sintesi</h2>`n<table><thead><tr><th>Problema</th><th class=`"num`">PC</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
    }

    # PC per PC: i fatti di ogni PC, prima i PC messi peggio
    $sezionePc = ''
    $conProblemi = @($perPc | Where-Object { $_.Urgente.Count + $_.DaControllare.Count -gt 0 })
    if ($conProblemi.Count -gt 0) {
        $righe = foreach ($x in $conProblemi) {
            $fatti = @($x.Urgente | ForEach-Object { '<li><i class="punto c-critico"></i>{0}</li>' -f (& $h $_) }) + @($x.DaControllare | ForEach-Object { '<li><i class="punto c-attenzione"></i>{0}</li>' -f (& $h $_) })
            $nome = '<b>{0}</b>' -f (& $h $x.Nome)
            if ($x.Utilizzatore) { $nome += '<div class="nota">{0}</div>' -f (& $h $x.Utilizzatore) }
            '<tr><td class="controllo">{0}</td><td><ul class="fatti">{1}</ul></td></tr>' -f $nome, ($fatti -join '')
        }
        $sezionePc = "<h2>PC per PC</h2>`n<table class=`"tabella-larga`"><thead><tr><th>PC</th><th>Problemi</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
        $senza = @($perPc | Where-Object { $_.Urgente.Count + $_.DaControllare.Count -eq 0 } | ForEach-Object { $_.Nome } | Sort-Object)
        if ($senza.Count -gt 0) { $sezionePc += "`n" + ('<p class="nota">Senza problemi: {0}.</p>' -f (& $h ($senza -join ', '))) }
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
    $recenti = @($Cambiamenti | Where-Object { $_ -and $_.Cambiamento } | ForEach-Object {
            $d = ConvertFrom-HDDataTesto $_.Data
            if ($d -and $d -ge $limite) { [pscustomobject]@{ Data = $d; Voce = $_ } }
        } | Sort-Object Data -Descending | Select-Object -First 60)
    $sezioneCambiamenti = ''
    if ($recenti.Count -gt 0) {
        $righe = foreach ($r in $recenti) {
            '<tr><td class="data">{0}</td><td class="controllo">{1}</td><td>{2}</td></tr>' -f $r.Data.ToString('dd/MM/yyyy', $script:HDInvariante), (& $h $r.Voce.Hostname), (& $h $r.Voce.Cambiamento)
        }
        $sezioneCambiamenti = "<h2>Cambiamenti degli ultimi $GiorniCambiamenti giorni</h2>`n<table class=`"tabella-larga`"><thead><tr><th>Data</th><th>PC</th><th>Cosa è cambiato</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
    }

    # Elenco dei PC
    $sezioneElenco = ''
    if ($elencoPc.Count -gt 0) {
        $righe = foreach ($x in @($elencoPc | Sort-Object Nome)) {
            $s = 'OK'
            if ($stato.ContainsKey($x.Nome)) { $s = $stato[$x.Nome] }
            $sistema = (([string]$x.Valori['OS']) -replace '^(?i)Microsoft\s+', '').Trim()
            $versione = (([string]$x.Valori['OSVersion']) -replace '\s*\(.*$', '').Trim()
            if ($versione) { $sistema = "$sistema $versione" }
            $ram = ''
            $gb = Get-HDMemoriaGB $x.Valori['Memory']
            if ($gb -gt 0) { $ram = '{0} GB' -f [math]::Round($gb, 1) }
            $modello = Join-HDProduttoreModello $x.Valori['Manufacturer'] $x.Valori['Model']
            '<tr><td class="controllo"><i class="punto c-{0}"></i>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td class="num">{5}</td></tr>' -f $classi[$s], (& $h $x.Nome), (& $h $x.Utilizzatore), (& $h $modello), (& $h $sistema), (& $h $ram)
        }
        $sezioneElenco = "<h2>Elenco dei PC</h2>`n<table class=`"tabella-larga`"><thead><tr><th>PC</th><th>Utilizzatore</th><th>Modello</th><th>Sistema</th><th class=`"num`">RAM</th></tr></thead><tbody>`n" + (@($righe) -join "`n") + "`n</tbody></table>"
    }

    $titolo = $Cliente
    if (-not $titolo) { $titolo = 'Parco PC' }
    $sottotitolo = Format-HDQuantita $elencoPc.Count 'PC censito' 'PC censiti'
    $giorno = Format-HDData $Adesso 'dd/MM/yyyy'
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
$sezioneSintesi
$sezionePc
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
    Legge l'Excel del censimento e mette in una pagina i problemi in sintesi (quanti PC per ciascuno), i
    problemi PC per PC (es. "BitLocker non attivo", "Desktop remoto (RDP) attivo"), il parco PC in numeri,
    i cambiamenti degli ultimi 90 giorni (foglio Storico) e l'elenco dei PC. Gli stessi problemi sono nel
    foglio "Riepilogo" dell'Excel, ricalcolato a ogni censimento. Il file va nella cartella Report.

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
        $cambiamenti = @($tabella.Storico.Righe | ForEach-Object { [pscustomobject]@{ Data = $_.Valori['Data']; Hostname = $_.Valori['Hostname']; Cambiamento = $_.Valori['Cambiamento'] } })
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

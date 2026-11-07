#Requires -Version 5.1
<#
.SYNOPSIS
Menu del Toolkit Help Desk. Avvialo con doppio clic su Avvia-Toolkit.cmd.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Clear-Host è una funzione predefinita anche in Windows PowerShell 5.1')]
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$radice = $PSScriptRoot
Import-Module (Join-Path (Join-Path $radice 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
$modulo = Get-Module HelpDeskToolkit

function Show-Titolo {
    Clear-Host
    $admin = & $modulo { Test-HDAmministratore }
    Write-Host ''
    Write-Host '  ================================================' -ForegroundColor DarkCyan
    Write-Host '     TOOLKIT HELP DESK' -ForegroundColor Cyan
    Write-Host ('     PC: {0}' -f (& $modulo { Get-HDNomePC })) -ForegroundColor DarkCyan
    Write-Host '  ================================================' -ForegroundColor DarkCyan
    if ($admin) { Write-Host '     Avviato come amministratore' -ForegroundColor Green }
    else { Write-Host '     Utente normale: alcuni controlli sono limitati (opzione A)' -ForegroundColor DarkYellow }
    Write-Host ''
}

function Read-SiNo {
    param([string]$Domanda, [bool]$Predefinito = $false)
    $suffisso = ' [s/N]'
    if ($Predefinito) { $suffisso = ' [S/n]' }
    $r = Read-Host ($Domanda + $suffisso)
    if (-not $r) { return $Predefinito }
    $r -match '^(s|si|sì|y|yes)$'
}

function Show-Tabella {
    param($Oggetti, [string[]]$Proprieta)
    if (@($Oggetti).Count -eq 0) { Write-Host '  (nessun risultato)' -ForegroundColor DarkGray; return }
    $Oggetti | Format-Table -Property $Proprieta -AutoSize -Wrap | Out-String -Width 160 | Write-Host
}

function Show-Volumi {
    param($Volumi)
    $colori = @{ OK = 'Green'; ATTENZIONE = 'Yellow'; CRITICO = 'Red' }
    foreach ($v in $Volumi) {
        $barra = '#' * [int][math]::Round((100 - $v.LiberoPct) / 5)
        $riga = '  {0,-4} {1,8} GB liberi su {2,8} GB  ({3,5}%)  [{4,-20}]  {5}' -f $v.Unita, $v.LiberoGB, $v.TotaleGB, $v.LiberoPct, $barra, $v.Stato
        Write-Host $riga -ForegroundColor $colori[$v.Stato]
    }
}

do {
    Show-Titolo
    Write-Host '  CENSIMENTO PC (Excel)' -ForegroundColor White
    Write-Host '  1  Censisci questo PC'
    Write-Host '  2  Scansiona la rete e aggiungi i dispositivi all''Excel'
    Write-Host '  3  Importa un file di Advanced IP Scanner'
    Write-Host '  4  Nuovo file di censimento'
    Write-Host '  5  Apri il censimento in Excel'
    Write-Host '  6  Inserisci le rilevazioni rimaste in sospeso'
    Write-Host ''
    Write-Host '  VERIFICHE' -ForegroundColor White
    Write-Host '  7  Verifica rapida della postazione (prima dell''intervento)'
    Write-Host '  8  Spazio disco (e cosa si può liberare)'
    Write-Host '  9  Apri la cartella dei report'
    Write-Host ''
    Write-Host '  A  Esegui come amministratore'
    Write-Host '  Q  Esci'
    Write-Host ''
    $scelta = (Read-Host '  Scelta').Trim().ToUpper()
    try {
        switch ($scelta) {
            '1' { [void](Add-PCCensimento) }
            '2' {
                $rete = (Read-Host '  Rete da scansionare (INVIO = quella di questo PC, es. 192.168.1.0/24)').Trim()
                $p = @{ NelCensimento = $true; Avanzamento = { param($t) Write-Host "  $t" -ForegroundColor DarkGray } }
                if ($rete) { $p.Rete = $rete }
                $s = Invoke-ScansioneRete @p
                Show-Tabella $s.Dispositivi @('IP', 'Hostname', 'MAC', 'Gruppo', 'Fonte')
                if ($s.Import) { & $modulo { param($e) Show-HDEsitoImport $e } $s.Import }
            }
            '3' {
                $file = & $modulo { Select-HDFileScansione }
                if ($file) {
                    $esito = Import-ScansioneRete -Percorso $file
                    & $modulo { param($e) Show-HDEsitoImport $e } $esito
                }
            }
            '4' {
                $cliente = (Read-Host '  Nome del cliente').Trim()
                $f = New-Censimento -Cliente $cliente
                Write-Host "  Creato: $($f.FullName)" -ForegroundColor Green
            }
            '5' {
                $f = & $modulo { Resolve-HDFileCensimento }
                Invoke-Item -LiteralPath $f
            }
            '6' {
                $s = Sync-Censimento
                Write-Host ('  Rilevazioni inserite: {0}  ({1})' -f $s.Inserite, $s.Censimento) -ForegroundColor Green
            }
            '7' {
                & $modulo { $script:HDUltimaVerifica = $null }
                Test-Postazione
                $report = & $modulo { $script:HDUltimaVerifica }
                if ($report -and (Read-SiNo '  Aprire il report HTML?')) { Invoke-Item -LiteralPath $report }
            }
            '8' {
                Write-Host ''
                Show-Volumi (Get-SpazioDisco)
                Write-Host ''
                if (Read-SiNo '  Stimare lo spazio recuperabile (Temp, Windows Update, Cestino...)?' $true) {
                    Write-Host '  Calcolo in corso, può richiedere qualche minuto...' -ForegroundColor DarkGray
                    Show-Tabella (Get-StimaPulizia) @('Voce', 'Dimensione', 'Come', 'Nota')
                }
                $percorso = Read-Host '  Cartella di cui vedere le sottocartelle più grandi (INVIO = salta, es. C:\Users)'
                if ($percorso) { Show-Tabella (Get-CartelleGrandi -Percorso $percorso.Trim('"')) @('Cartella', 'Dimensione', 'File', 'NonLeggibili') }
                Write-Host '  Nota: il toolkit non cancella nulla. Pulisci a mano con Pulizia disco, dopo aver chiesto all''utente.' -ForegroundColor DarkGray
            }
            '9' { Invoke-Item -LiteralPath (& $modulo { Get-HDCartella -Tipo Report }) }
            'A' {
                $exe = (Get-Process -Id $PID).Path
                Start-Process -FilePath $exe -Verb RunAs -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
                $scelta = 'Q'
            }
            'Q' { }
            default { Write-Host '  Scelta non valida' -ForegroundColor Yellow }
        }
    }
    catch {
        Write-Host ''
        Write-Host "  ERRORE: $($_.Exception.Message)" -ForegroundColor Red
    }
    if ($scelta -ne 'Q') { [void](Read-Host "`n  Premi INVIO per tornare al menu") }
} while ($scelta -ne 'Q')

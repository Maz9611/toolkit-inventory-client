#Requires -Version 5.1
<#
.SYNOPSIS
Importa nel censimento l'export di Advanced IP Scanner (usato da Importa-Scansione.cmd).
#>
[CmdletBinding()]
param([Parameter(Position = 0)][string]$Percorso, [string]$Cliente)

$ErrorActionPreference = 'Stop'
try {
    Import-Module (Join-Path (Join-Path $PSScriptRoot 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
    $modulo = Get-Module HelpDeskToolkit
    if (-not $Percorso) { $Percorso = & $modulo { Select-HDFileScansione } }
    if (-not $Percorso) { return }
    $esito = Import-ScansioneRete -Percorso $Percorso -Cliente $Cliente
    & $modulo { param($e) Show-HDEsitoImport $e } $esito
}
catch {
    Write-Host ''
    Write-Host ('ERRORE: {0}' -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}

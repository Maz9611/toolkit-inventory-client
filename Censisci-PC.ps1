#Requires -Version 5.1
<#
.SYNOPSIS
Censisce questo PC nell'Excel del censimento (usato da Censisci-PC.cmd).
#>
[CmdletBinding()]
param([string]$Censimento)

$ErrorActionPreference = 'Stop'
try {
    Import-Module (Join-Path (Join-Path $PSScriptRoot 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
    $esito = Add-PCCensimento -Censimento $Censimento
    if ($esito.Scritto) { Write-Host ''; Write-Host ('Fatto. Excel: {0}' -f $esito.Censimento) -ForegroundColor Green }
}
catch {
    Write-Host ''
    Write-Host ('ERRORE: {0}' -f $_.Exception.Message) -ForegroundColor Red
    exit 1
}

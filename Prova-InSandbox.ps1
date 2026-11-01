#Requires -Version 5.1
<#
.SYNOPSIS
Apre una Windows Sandbox (Windows pulito e usa-e-getta) con il toolkit già dentro, per provarlo.

.DESCRIPTION
Crea un file .wsb che collega questa cartella alla Sandbox in SOLA LETTURA e la avvia.
Nella Sandbox: apri la cartella "helpdesk-toolkit" sul Desktop e fai doppio clic su
Avvia-Prova.cmd (poi prova anche il menu con Avvia-Toolkit.cmd).
Alla chiusura della Sandbox tutto viene cancellato.

Requisiti: Windows 10/11 Pro, Enterprise o Education con la funzionalità
"Windows Sandbox" attiva (Pannello di controllo > Programmi > Attiva o disattiva
funzionalità di Windows). Non sostituisce la prova su un PC simile a quelli dei
clienti (con dominio, antivirus, stampanti...).
#>
[CmdletBinding()]
param()

$sandbox = Join-Path $env:windir 'System32\WindowsSandbox.exe'
if (-not (Test-Path -LiteralPath $sandbox)) {
    throw 'Windows Sandbox non è attiva su questo PC: attivala da "Attiva o disattiva funzionalità di Windows" (serve Windows Pro/Enterprise) e riavvia.'
}
$cartella = $PSScriptRoot
$xml = @"
<Configuration>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$([System.Security.SecurityElement]::Escape($cartella))</HostFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <Networking>Default</Networking>
</Configuration>
"@
$file = Join-Path ([System.IO.Path]::GetTempPath()) 'HelpDeskToolkit-Prova.wsb'
[System.IO.File]::WriteAllText($file, $xml, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Configurazione Sandbox: $file"
Write-Host "Nella Sandbox apri la cartella '$(Split-Path -Leaf $cartella)' sul Desktop e lancia Avvia-Prova.cmd" -ForegroundColor Cyan
Start-Process -FilePath $file

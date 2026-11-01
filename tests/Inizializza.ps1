# Inizializzazione comune ai test: da richiamare con  . $PSScriptRoot\Inizializza.ps1  dentro BeforeAll.
# Importa il modulo e lo fa lavorare con una configurazione di prova dentro $TestDrive.
$script:RadiceToolkit = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path (Join-Path $script:RadiceToolkit 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -Force
$script:ConfigPrecedente = $env:HELPDESK_TOOLKIT_CONFIG
$script:CartellaReportTest = Join-Path $TestDrive 'Report'
$script:CartellaCensimentoTest = Join-Path $TestDrive 'Censimento'
$percorsoConfig = Join-Path $TestDrive 'config-test.psd1'
$testoConfig = @"
@{
    Azienda        = 'Azienda Test'
    Tecnico        = 'Tecnico Test'
    CartellaReport = '$($script:CartellaReportTest.Replace("'", "''"))'
    Censimento     = @{ Cartella = '$($script:CartellaCensimentoTest.Replace("'", "''"))' }
}
"@
[System.IO.File]::WriteAllText($percorsoConfig, $testoConfig, (New-Object System.Text.UTF8Encoding($true)))
$env:HELPDESK_TOOLKIT_CONFIG = $percorsoConfig

# HelpDeskToolkit - modulo principale.
# Carica le funzioni private (Private\*.ps1) e pubbliche (Public\*.ps1).

$script:HDModuloRadice = $PSScriptRoot
$script:HDToolkitRadice = Split-Path -Parent $PSScriptRoot
$script:HDTestoProduttoriMac = $null   # nell'exe arriva dall'avvio, qui si legge da Dati\produttori-mac.txt.gz
$script:HDProduttoriMac = $null

foreach ($sottocartella in 'Private', 'Public') {
    $percorso = Join-Path $PSScriptRoot $sottocartella
    foreach ($file in (Get-ChildItem -Path $percorso -Filter '*.ps1' -File | Sort-Object Name)) {
        . $file.FullName
    }
}

Export-ModuleMember -Function @(
    'Get-SpazioDisco'
    'Get-CartelleGrandi'
    'Get-StimaPulizia'
    'Test-Postazione'
    'Get-InventarioPC'
    'New-Censimento'
    'Import-ScansioneRete'
    'Add-PCCensimento'
    'Sync-Censimento'
    'Invoke-ScansioneRete'
    'Export-RiepilogoCensimento'
)

@{
    RootModule           = 'HelpDeskToolkit.psm1'
    ModuleVersion        = '2.4.0'
    GUID                 = '3f6b2c1e-8a4d-4f7b-9c2e-5d1a7e9b0c42'
    Author               = 'Toolkit Help Desk contributors'
    Description          = 'Toolkit help desk: censimento PC su Excel con scansione della rete, verifica postazione e spazio disco.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @(
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
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags = @('HelpDesk', 'Inventario', 'Censimento', 'Excel', 'Windows')
        }
    }
}

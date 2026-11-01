BeforeAll { . (Join-Path $PSScriptRoot 'Inizializza.ps1') }
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Date degli aggiornamenti (Win32_QuickFixEngineering)' {
    It '"<Valore>" -> <Atteso>' -TestCases @(
        @{ Valore = '10/15/2026'; Atteso = '2026-10-15' }
        @{ Valore = '25/10/2026'; Atteso = '2026-10-25' }
        @{ Valore = '01d93b5a7c4e8000'; Atteso = 'filetime' }
        @{ Valore = ''; Atteso = $null }
        @{ Valore = '11/12/2026'; Atteso = $null }
    ) {
        $d = InModuleScope HelpDeskToolkit -Parameters @{ V = $Valore } { param($V) ConvertFrom-HDDataQfe -Valore $V -Adesso ([datetime]'2026-11-05') }
        if ($null -eq $Atteso) { $d | Should -BeNullOrEmpty }
        elseif ($Atteso -eq 'filetime') { $d | Should -BeOfType [datetime]; $d.Year | Should -Be 2023 }
        else { $d | Should -Be ([datetime]$Atteso) }
    }
}

Describe 'Supporto del sistema operativo' {
    It '<Caption> -> <Atteso>' -TestCases @(
        @{ Caption = 'Windows 11 Pro'; Atteso = 'OK' }
        @{ Caption = 'Windows 10 Pro'; Atteso = 'ATTENZIONE' }
        @{ Caption = 'Windows 10 Enterprise LTSC 2021'; Atteso = 'INFO' }
        @{ Caption = 'Windows 7 Professional'; Atteso = 'CRITICO' }
        @{ Caption = 'Windows 8.1 Pro'; Atteso = 'CRITICO' }
        @{ Caption = 'Windows Server 2019 Standard'; Atteso = 'INFO' }
    ) {
        (InModuleScope HelpDeskToolkit -Parameters @{ C = $Caption } { param($C) Get-HDSupportoOS -Caption $C }).Stato | Should -Be $Atteso
    }
}

Describe 'Numero di serie e antivirus' {
    It 'scarta i numeri di serie segnaposto' {
        InModuleScope HelpDeskToolkit {
            Get-HDNumeroSerie 'To be filled by O.E.M.' | Should -Be ''
            Get-HDNumeroSerie 'Default string' | Should -Be ''
            Get-HDNumeroSerie '  CZC1234  ' | Should -Be 'CZC1234'
        }
    }

    It 'decodifica lo stato dell''antivirus (SecurityCenter2)' {
        InModuleScope HelpDeskToolkit {
            $a = ConvertFrom-HDStatoAntivirus 397568     # 0x061100
            $a.Attivo | Should -BeTrue
            $a.Aggiornato | Should -BeTrue
            (ConvertFrom-HDStatoAntivirus 393472).Attivo | Should -BeFalse          # 0x060100
            (ConvertFrom-HDStatoAntivirus 397584).Aggiornato | Should -BeFalse      # 0x061110
        }
    }
}

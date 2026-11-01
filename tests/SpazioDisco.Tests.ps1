BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    function New-DiscoFinto([string]$Unita, [double]$TotaleGB, [double]$LiberoGB) {
        [pscustomobject]@{ DeviceID = $Unita; VolumeName = 'Dati'; FileSystem = 'NTFS'; Size = $TotaleGB * 1GB; FreeSpace = $LiberoGB * 1GB }
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Soglie dello spazio libero' {
    It '<LiberoGB> GB liberi (<Pct>%) -> <Atteso>' -TestCases @(
        @{ LiberoGB = 100; Pct = 50; Atteso = 'OK' }
        @{ LiberoGB = 15; Pct = 30; Atteso = 'ATTENZIONE' }
        @{ LiberoGB = 50; Pct = 8; Atteso = 'ATTENZIONE' }
        @{ LiberoGB = 8; Pct = 40; Atteso = 'CRITICO' }
        @{ LiberoGB = 200; Pct = 4; Atteso = 'CRITICO' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ G = $LiberoGB; P = $Pct } { param($G, $P) Get-HDStatoSpazio -LiberoGB $G -LiberoPct $P } | Should -Be $Atteso
    }
}

Describe 'Get-SpazioDisco' {
    It 'calcola GB e percentuali e ignora i dischi senza dimensione (lettori vuoti)' {
        Mock -ModuleName HelpDeskToolkit Get-HDCim {
            New-DiscoFinto 'C:' 476.9 120.4
            New-DiscoFinto 'D:' 931.5 3
            [pscustomobject]@{ DeviceID = 'E:'; Size = $null; FreeSpace = $null }
        }
        $v = @(Get-SpazioDisco)
        $v | Should -HaveCount 2
        $v[0].Unita | Should -Be 'C:'
        $v[0].LiberoGB | Should -Be 120.4
        $v[0].LiberoPct | Should -Be 25.2
        $v[0].Stato | Should -Be 'OK'
        $v[1].Stato | Should -Be 'CRITICO'
    }

    It 'chiede solo i dischi fissi, o anche i rimovibili con -IncludiRimovibili' {
        Mock -ModuleName HelpDeskToolkit Get-HDCim { }
        Get-SpazioDisco
        Should -Invoke -ModuleName HelpDeskToolkit Get-HDCim -Times 1 -ParameterFilter { $Filtro -eq 'DriveType=3' }
        Get-SpazioDisco -IncludiRimovibili
        Should -Invoke -ModuleName HelpDeskToolkit Get-HDCim -Times 1 -ParameterFilter { $Filtro -eq 'DriveType=2 OR DriveType=3' }
    }

    It 'se un PC non risponde avvisa e prosegue con gli altri' {
        Mock -ModuleName HelpDeskToolkit Get-HDCim { throw 'RPC non disponibile' } -ParameterFilter { $ComputerName -eq 'PC-SPENTO' }
        Mock -ModuleName HelpDeskToolkit Get-HDCim { New-DiscoFinto 'C:' 100 50 } -ParameterFilter { $ComputerName -eq 'PC-OK' }
        $v = @(Get-SpazioDisco -ComputerName 'PC-SPENTO', 'PC-OK' -WarningVariable avvisi -WarningAction SilentlyContinue)
        $v | Should -HaveCount 1
        $v[0].Computer | Should -Be 'PC-OK'
        "$avvisi" | Should -Match 'PC-SPENTO'
    }
}

Describe 'Dimensione delle cartelle' {
    BeforeAll {
        $base = Join-Path $TestDrive 'albero'
        $grande = Join-Path $base 'grande'
        $sotto = Join-Path $grande 'sotto'
        $piccola = Join-Path $base 'piccola'
        foreach ($d in $sotto, $piccola, (Join-Path $base 'vuota')) { [void][System.IO.Directory]::CreateDirectory($d) }
        [System.IO.File]::WriteAllBytes((Join-Path $grande 'a.bin'), (New-Object byte[] 3000))
        [System.IO.File]::WriteAllBytes((Join-Path $sotto 'b.bin'), (New-Object byte[] 2000))
        [System.IO.File]::WriteAllBytes((Join-Path $piccola 'c.bin'), (New-Object byte[] 100))
        [System.IO.File]::WriteAllBytes((Join-Path $base 'sciolto.bin'), (New-Object byte[] 50))
    }

    It 'Measure-HDCartella somma i file di tutte le sottocartelle' {
        $m = InModuleScope HelpDeskToolkit -Parameters @{ P = (Join-Path $base 'grande') } { param($P) Measure-HDCartella -Percorso $P }
        $m.Byte | Should -Be 5000
        $m.File | Should -Be 2
        $m.Errori | Should -Be 0
    }

    It 'Get-CartelleGrandi ordina dalla più grande e rispetta -Primi' {
        $c = @(Get-CartelleGrandi -Percorso $base)
        $c[0].Cartella | Should -BeLike '*grande'
        $c[0].Byte | Should -Be 5000
        ($c | Where-Object { $_.Cartella -like '*(file nella cartella)' }).Byte | Should -Be 50
        @(Get-CartelleGrandi -Percorso $base -Primi 2) | Should -HaveCount 2
    }

    It 'Get-CartelleGrandi segnala una cartella inesistente' {
        { Get-CartelleGrandi -Percorso (Join-Path $TestDrive 'non-esiste') } | Should -Throw '*non trovata*'
    }

    It 'Get-StimaPulizia misura cartelle e file e salta i percorsi che non esistono' {
        $dump = Join-Path $base 'MEMORY.DMP'
        [System.IO.File]::WriteAllBytes($dump, (New-Object byte[] 1234))
        Mock -ModuleName HelpDeskToolkit Get-HDVociPulizia {
            @{ Voce = 'Temp'; Percorso = (Join-Path $base 'grande'); Come = 'x' }
            @{ Voce = 'Dump'; Percorso = $dump; Come = 'y' }
            @{ Voce = 'Assente'; Percorso = (Join-Path $base 'manca'); Come = 'z' }
        }
        $s = @(Get-StimaPulizia)
        $s | Should -HaveCount 2
        ($s | Where-Object Voce -eq 'Temp').Byte | Should -Be 5000
        ($s | Where-Object Voce -eq 'Dump').Byte | Should -Be 1234
    }
}

BeforeAll { . (Join-Path $PSScriptRoot 'Inizializza.ps1') }
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Formattazione all''italiana' {
    It 'Format-HDByte: <Byte> -> <Atteso>' -TestCases @(
        @{ Byte = 1610612736; Atteso = '1,5 GB' }
        @{ Byte = 524288000; Atteso = '500 MB' }
        @{ Byte = 2199023255552; Atteso = '2,00 TB' }
        @{ Byte = 512; Atteso = '512 byte' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ B = $Byte } { param($B) Format-HDByte $B } | Should -Be $Atteso
    }

    It 'ConvertTo-HDValoreTesto usa virgola decimale, date gg/mm/aaaa e Sì/No' {
        InModuleScope HelpDeskToolkit {
            ConvertTo-HDValoreTesto 12.5 | Should -Be '12,5'
            ConvertTo-HDValoreTesto ([datetime]'2026-11-01') | Should -Be '01/11/2026'
            ConvertTo-HDValoreTesto ([datetime]'2026-11-01 09:05') | Should -Be '01/11/2026 09:05'
            ConvertTo-HDValoreTesto $true | Should -Be 'Sì'
            ConvertTo-HDValoreTesto @('a', 'b') | Should -Be 'a, b'
            ConvertTo-HDValoreTesto $null | Should -Be ''
        }
    }

    It 'ConvertTo-HDNomeFile toglie accenti e simboli' {
        InModuleScope HelpDeskToolkit {
            ConvertTo-HDNomeFile 'Studio Rossi & Figli s.r.l. – Città' | Should -Be 'Studio-Rossi-Figli-s-r-l-Citta'
            ConvertTo-HDNomeFile '' | Should -Be 'senza-nome'
            (ConvertTo-HDNomeFile ('x' * 100)).Length | Should -BeLessOrEqual 40
        }
    }

    It 'singolare/plurale e produttore non ripetuto' {
        InModuleScope HelpDeskToolkit {
            Format-HDQuantita 1 'stampante' 'stampanti' | Should -Be '1 stampante'
            Format-HDQuantita 3 'stampante' 'stampanti' | Should -Be '3 stampanti'
            Join-HDProduttoreModello 'HP' 'HP ProDesk 400 G7' | Should -Be 'HP ProDesk 400 G7'
            Join-HDProduttoreModello 'Dell Inc.' 'OptiPlex 7090' | Should -Be 'Dell Inc. OptiPlex 7090'
            Join-HDProduttoreModello 'LENOVO' '20XW00GBIX' | Should -Be 'LENOVO 20XW00GBIX'
        }
    }

    It 'ConvertTo-HDChiave normalizza maiuscole, accenti e spazi' {
        InModuleScope HelpDeskToolkit { ConvertTo-HDChiave '  Modalità   Intervento ' } | Should -Be 'modalita intervento'
    }

    It 'ConvertTo-HDHtml impedisce di iniettare codice nella pagina' {
        InModuleScope HelpDeskToolkit { ConvertTo-HDHtml '<script>alert("x")</script> & co' } |
            Should -Be '&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt; &amp; co'
    }
}

Describe 'Lettura dei file di testo in qualsiasi codifica' {
    BeforeAll { $testo = "Cliente: Città perché €`r`nRiga 2" }

    It 'legge <Nome>' -TestCases @(
        @{ Nome = 'UTF-8 con BOM'; Codifica = 'utf8bom' }
        @{ Nome = 'UTF-8 senza BOM (Blocco note recente)'; Codifica = 'utf8' }
        @{ Nome = 'ANSI Windows-1252 (Blocco note vecchio)'; Codifica = '1252' }
        @{ Nome = 'UTF-16 LE ("Unicode" di Blocco note)'; Codifica = 'utf16' }
    ) {
        $enc = switch ($Codifica) {
            'utf8bom' { New-Object System.Text.UTF8Encoding($true) }
            'utf8' { New-Object System.Text.UTF8Encoding($false) }
            '1252' { [System.Text.Encoding]::GetEncoding(1252) }
            'utf16' { [System.Text.Encoding]::Unicode }
        }
        $p = Join-Path $TestDrive "codifica-$Codifica.txt"
        [System.IO.File]::WriteAllText($p, $testo, $enc)
        InModuleScope HelpDeskToolkit -Parameters @{ P = $p } { param($P) Read-HDFileTesto $P } | Should -BeExactly $testo
    }

    It 'Write-HDFileTesto scrive UTF-8 con BOM' {
        $p = Join-Path $TestDrive 'scritto.txt'
        InModuleScope HelpDeskToolkit -Parameters @{ P = $p } { param($P) Write-HDFileTesto -Percorso $P -Testo 'àè' }
        $b = [System.IO.File]::ReadAllBytes($p)
        $b[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
    }
}

Describe 'Configurazione e cartelle' {
    It 'unisce config.psd1 ai valori predefiniti' {
        $cfg = InModuleScope HelpDeskToolkit { Get-HDConfig }
        $cfg.Azienda | Should -Be 'Azienda Test'
        $cfg.Tecnico | Should -Be 'Tecnico Test'
        $cfg.Soglie.UptimeGiorni | Should -Be 14
        $cfg.Censimento.RilevaCynet | Should -Be 'Cynet'
    }

    It 'con un config.psd1 rovinato avvisa e usa i predefiniti' {
        $p = Join-Path $TestDrive 'rotto.psd1'
        Set-Content -LiteralPath $p -Value '@{ Azienda = ' -Encoding UTF8
        $cfg = InModuleScope HelpDeskToolkit -Parameters @{ P = $p } { param($P) Get-HDConfig -Percorso $P -WarningAction SilentlyContinue }
        $cfg.Azienda | Should -Be 'La tua azienda'
    }

    It 'Get-HDCartella crea la cartella indicata in config' {
        $c = InModuleScope HelpDeskToolkit { Get-HDCartella -Tipo Report }
        $c | Should -Be $script:CartellaReportTest
        Test-Path -LiteralPath $c | Should -BeTrue
    }

    It 'Get-HDPercorsoUnico non sovrascrive i file esistenti' {
        New-Item -ItemType File -Path (Join-Path $TestDrive 'doppio.html') -Force | Out-Null
        $p = InModuleScope HelpDeskToolkit -Parameters @{ C = $TestDrive } { param($C) Get-HDPercorsoUnico -Cartella $C -NomeBase 'doppio' -Estensione '.html' }
        Split-Path -Leaf $p | Should -Be 'doppio-2.html'
    }
}

Describe 'Modulo' {
    It 'il manifest è valido' {
        { Test-ModuleManifest -Path (Join-Path (Join-Path $script:RadiceToolkit 'HelpDeskToolkit') 'HelpDeskToolkit.psd1') -ErrorAction Stop } | Should -Not -Throw
    }

    It 'esporta solo i comandi pubblici' {
        $attesi = 'Add-PCCensimento', 'Get-CartelleGrandi', 'Get-InventarioPC', 'Get-SpazioDisco', 'Get-StimaPulizia',
        'Import-ScansioneRete', 'Invoke-ScansioneRete', 'New-Censimento', 'Sync-Censimento', 'Test-Postazione'
        $funzioni = @(Get-Command -Module HelpDeskToolkit -CommandType Function | ForEach-Object { $_.Name } | Sort-Object)
        $funzioni | Should -Be ($attesi | Sort-Object)
    }

    It 'ogni comando pubblico ha la guida (Get-Help) in italiano' {
        foreach ($c in Get-Command -Module HelpDeskToolkit -CommandType Function) {
            (Get-Help $c.Name).Synopsis | Should -Not -BeLike "$($c.Name)*" -Because "$($c.Name) deve avere .SYNOPSIS"
        }
    }
}

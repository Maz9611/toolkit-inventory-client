BeforeDiscovery {
    $radice = Split-Path -Parent $PSScriptRoot
    $script:FilePowerShell = @(Get-ChildItem -Path $radice -Recurse -File -Include '*.ps1', '*.psm1', '*.psd1' |
            ForEach-Object { @{ Nome = $_.FullName.Substring($radice.Length + 1); Percorso = $_.FullName } })
    $script:FileCmd = @(Get-ChildItem -Path $radice -File -Filter '*.cmd' | ForEach-Object { @{ Nome = $_.Name; Percorso = $_.FullName } })
}

Describe 'Qualità dei file PowerShell' {
    It '<Nome> non ha errori di sintassi' -TestCases $script:FilePowerShell {
        $token = $null
        $errori = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($Percorso, [ref]$token, [ref]$errori)
        $errori | Should -BeNullOrEmpty
    }

    It '<Nome> ha il BOM UTF-8 se contiene lettere accentate (serve a Windows PowerShell 5.1)' -TestCases $script:FilePowerShell {
        $byte = [System.IO.File]::ReadAllBytes($Percorso)
        $soloAscii = -not ($byte | Where-Object { $_ -gt 127 } | Select-Object -First 1)
        $haBom = $byte.Length -ge 3 -and $byte[0] -eq 0xEF -and $byte[1] -eq 0xBB -and $byte[2] -eq 0xBF
        ($soloAscii -or $haBom) | Should -BeTrue -Because 'senza BOM PowerShell 5.1 legge gli accenti come caratteri strani'
    }

    It '<Nome> non contiene virgolette "tipografiche" (PowerShell le tratta come delimitatori)' -TestCases $script:FilePowerShell {
        $testo = [System.IO.File]::ReadAllText($Percorso)
        $testo.IndexOfAny([char[]]@(0x2018, 0x2019, 0x201A, 0x201B, 0x201C, 0x201D, 0x201E)) | Should -Be -1
    }
}

Describe 'Launcher .cmd' {
    It '<Nome> è solo ASCII e con a capo Windows (CRLF)' -TestCases $script:FileCmd {
        $byte = [System.IO.File]::ReadAllBytes($Percorso)
        @($byte | Where-Object { $_ -gt 127 }) | Should -HaveCount 0
        $testo = [System.Text.Encoding]::ASCII.GetString($byte)
        ($testo -replace "`r`n", '') | Should -Not -Match "`n"
    }
}

Describe 'PSScriptAnalyzer (compatibilità con Windows PowerShell 5.1)' {
    It 'il codice del toolkit non ha avvisi' -Skip:(-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
        $radice = Split-Path -Parent $PSScriptRoot
        $impostazioni = Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1'
        $risultati = @(Get-ChildItem -Path $radice -Recurse -File -Include '*.ps1', '*.psm1', '*.psd1' |
                Where-Object { $_.DirectoryName -ne $PSScriptRoot } |
                ForEach-Object { Invoke-ScriptAnalyzer -Path $_.FullName -Settings $impostazioni })
        $testo = ($risultati | ForEach-Object { '{0}:{1} {2} {3}' -f $_.ScriptName, $_.Line, $_.RuleName, $_.Message }) -join "`n"
        $risultati | Should -HaveCount 0 -Because $testo
    }
}

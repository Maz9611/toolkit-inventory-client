BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    $script:CreaExe = Join-Path (Join-Path $script:RadiceToolkit 'exe') 'Crea-Exe.ps1'
    $script:Pacchetto = & $script:CreaExe -SoloPacchetto
    if (-not ('ToolkitHelpDesk.Esecutore' -as [type])) {
        # Lo stesso codice C# dell'exe, senza la parte Windows Forms: compila anche come C# 5
        $codice = [System.IO.File]::ReadAllText((Join-Path (Join-Path $script:RadiceToolkit 'exe') 'Avvio.cs'))
        $codice = $codice.Replace('@@VERSIONE@@', '0.0.0.1').Replace('@@AUTORE@@', 'test')
        Add-Type -TypeDefinition $codice -CompilerOptions '-define:SENZA_FINESTRA' -IgnoreWarnings
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Programma ToolkitHelpDesk.exe' {
    It 'Avvio.cs è solo ASCII (compila con qualsiasi codifica)' {
        $byte = [System.IO.File]::ReadAllBytes((Join-Path (Join-Path $script:RadiceToolkit 'exe') 'Avvio.cs'))
        @($byte | Where-Object { $_ -gt 127 }).Count | Should -Be 0
    }

    It 'il pacchetto contiene tutto il modulo e la finestra' {
        $manifesto = Import-PowerShellDataFile (Join-Path (Join-Path $script:RadiceToolkit 'HelpDeskToolkit') 'HelpDeskToolkit.psd1')
        $script:Pacchetto.Versione | Should -Be $manifesto.ModuleVersion
        foreach ($f in Get-ChildItem -Path (Join-Path $script:RadiceToolkit 'HelpDeskToolkit') -Recurse -Filter '*.ps1') {
            $script:Pacchetto.Modulo | Should -Match ([regex]::Escape('#region ' + $f.Directory.Name + '\' + $f.Name))
        }
        $script:Pacchetto.Interfaccia | Should -Not -Match '(?m)^#Requires'
        $script:Pacchetto.Interfaccia | Should -Match 'Censisci questo PC'
        { [scriptblock]::Create($script:Pacchetto.Modulo) } | Should -Not -Throw
        { [scriptblock]::Create($script:Pacchetto.Interfaccia) } | Should -Not -Throw
    }

    It 'il motore dell''exe carica il modulo, usa la cartella dell''exe e restituisce il codice di uscita' {
        $radiceExe = Join-Path $TestDrive 'chiavetta'
        [void][System.IO.Directory]::CreateDirectory($radiceExe)
        $esito = Join-Path $TestDrive 'esito-exe.txt'
        $finta = @'
[CmdletBinding()]
param([switch]$Verifica)
$ErrorActionPreference = 'Stop'
Write-Host 'senza console' -ForegroundColor Green
$m = Get-Module HelpDeskToolkit
$righe = @(
    'radice=' + (& $m { $script:HDToolkitRadice })
    'comandi=' + ((Get-Command -Module HelpDeskToolkit | ForEach-Object { $_.Name } | Sort-Object) -join ',')
    'interattivo=' + (& $m { Test-HDInterattivo })
    'verifica=' + $Verifica
    'exe=' + $HDPercorsoExe
)
[System.IO.File]::WriteAllLines('@@ESITO@@', [string[]]$righe)
exit 7
'@.Replace('@@ESITO@@', $esito.Replace("'", "''"))
        $codice = [ToolkitHelpDesk.Esecutore]::Esegui($script:Pacchetto.Modulo, $finta, $radiceExe, 'X:\ToolkitHelpDesk.exe', $true, $false)
        $codice | Should -Be 7
        $valori = @{}
        foreach ($r in [System.IO.File]::ReadAllLines($esito)) { $k, $v = $r -split '=', 2; $valori[$k] = $v }
        $valori['radice'] | Should -Be $radiceExe
        $valori['comandi'] | Should -Be ((@($script:Pacchetto.Funzioni) | Sort-Object) -join ',')
        $valori['interattivo'] | Should -Be 'False'
        $valori['verifica'] | Should -Be 'True'
        $valori['exe'] | Should -Be 'X:\ToolkitHelpDesk.exe'
    }

    It 'senza "exit" il codice di uscita è 0' {
        [ToolkitHelpDesk.Esecutore]::Esegui($script:Pacchetto.Modulo, 'param([switch]$Verifica) $null = Get-Module HelpDeskToolkit', $TestDrive, 'X:\a.exe', $false, $false) | Should -Be 0
    }
}

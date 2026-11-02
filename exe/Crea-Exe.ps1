#Requires -Version 5.1
<#
.SYNOPSIS
Crea ToolkitHelpDesk.exe: un solo file con dentro il modulo e la finestra del toolkit.

.DESCRIPTION
Comprime il modulo HelpDeskToolkit e Interfaccia.ps1 dentro un piccolo programma C#
(exe\Avvio.cs) che li esegue con il Windows PowerShell 5.1 di Windows.

- Con Windows PowerShell 5.1 (Crea-Exe.cmd) usa il compilatore C# di .NET Framework che c'è
  in ogni Windows: non scarica e non installa niente.
- Con PowerShell 7 usa il compilatore Roslyn incluso in PowerShell. Su Windows trova da solo
  gli assembly di .NET Framework; altrove serve -Riferimenti: una cartella con mscorlib.dll,
  System.dll, System.Core.dll, System.Drawing.dll, System.Windows.Forms.dll di .NET Framework 4.8
  e System.Management.Automation.dll di Windows PowerShell 5.1 (pacchetti NuGet di riferimento).

Dopo ogni modifica al modulo o alla finestra va ricreato, e provato con Avvia-Prova.cmd.

.PARAMETER Destinazione
Dove scrivere l'exe. Predefinito: ToolkitHelpDesk.exe nella cartella del toolkit.

.PARAMETER SoloPacchetto
Non compila: restituisce i testi che finirebbero nell'exe (per i test).

.EXAMPLE
.\exe\Crea-Exe.ps1
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleTypes', '', Justification = 'I tipi Roslyn si usano solo con PowerShell 7, dopo averli caricati')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Ramo PowerShell 7 separato dal ramo Windows PowerShell')]
[CmdletBinding()]
param(
    [string]$Destinazione,
    [string]$Riferimenti,
    [switch]$SoloPacchetto
)

$ErrorActionPreference = 'Stop'
$cartellaExe = $PSScriptRoot
$radice = Split-Path -Parent $cartellaExe
$cartellaModulo = Join-Path $radice 'HelpDeskToolkit'

function Read-TestoUtf8 {
    param([string]$Percorso)
    $testo = [System.IO.File]::ReadAllText($Percorso, (New-Object System.Text.UTF8Encoding($false)))
    $testo.TrimStart([char]0xFEFF)
}

function Get-PacchettoToolkit {
    # I due testi PowerShell che l'exe contiene: il modulo (come modulo dinamico) e la finestra
    $manifesto = Import-PowerShellDataFile (Join-Path $cartellaModulo 'HelpDeskToolkit.psd1')
    $parti = New-Object System.Collections.Generic.List[string]
    $parti.Add('param($Radice, $TestoProduttoriMac)')
    $parti.Add('# Modulo HelpDeskToolkit {0}, generato da exe\Crea-Exe.ps1: non modificare qui.' -f $manifesto.ModuleVersion)
    $parti.Add('$script:HDModuloRadice = $Radice')
    $parti.Add('$script:HDToolkitRadice = $Radice')
    $parti.Add('$script:HDTestoProduttoriMac = $TestoProduttoriMac')
    foreach ($sottocartella in 'Private', 'Public') {
        foreach ($file in (Get-ChildItem -Path (Join-Path $cartellaModulo $sottocartella) -Filter '*.ps1' -File | Sort-Object Name)) {
            $parti.Add(('#region {0}\{1}' -f $sottocartella, $file.Name))
            $parti.Add((Read-TestoUtf8 $file.FullName))
            $parti.Add('#endregion')
        }
    }
    $elenco = @($manifesto.FunctionsToExport | ForEach-Object { "'$_'" }) -join ', '
    $parti.Add(('Export-ModuleMember -Function @({0})' -f $elenco))

    $interfaccia = Read-TestoUtf8 (Join-Path $radice 'Interfaccia.ps1')
    $interfaccia = [regex]::Replace($interfaccia, '(?m)^#Requires.*$', '# (requisiti: Windows PowerShell 5.1, li garantisce l''exe)')

    [pscustomobject]@{
        Modulo      = ($parti -join "`r`n")
        Interfaccia = $interfaccia
        Versione    = [string]$manifesto.ModuleVersion
        Autore      = [string]$manifesto.Author
        Funzioni    = @($manifesto.FunctionsToExport)
    }
}

function ConvertTo-Gzip {
    param([string]$Testo)
    $memoria = New-Object System.IO.MemoryStream
    $gz = New-Object System.IO.Compression.GZipStream($memoria, [System.IO.Compression.CompressionMode]::Compress)
    $byte = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Testo)
    $gz.Write($byte, 0, $byte.Length)
    $gz.Dispose()
    , $memoria.ToArray()
}

function Get-RiferimentiNet48 {
    if ($Riferimenti) {
        $cartelle = @($Riferimenti)
    }
    else {
        $cartelle = @((Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'), (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'))
    }
    $nomi = 'mscorlib.dll', 'System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll'
    $trovati = foreach ($n in $nomi) {
        $f = $cartelle | ForEach-Object { Join-Path $_ $n } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $f) { throw "Non trovo $n di .NET Framework 4.8 (usa -Riferimenti)." }
        $f
    }
    $sma = $null
    if ($Riferimenti) { $sma = Join-Path $Riferimenti 'System.Management.Automation.dll' }
    else { $sma = Join-Path $env:WINDIR 'Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation\v4.0_3.0.0.0__31bf3856ad364e35\System.Management.Automation.dll' }
    if (-not (Test-Path -LiteralPath $sma)) { throw "Non trovo System.Management.Automation.dll di Windows PowerShell 5.1 ($sma)." }
    @($trovati) + @($sma)
}

function Invoke-CompilaConCodeDom {
    param([string]$Codice, [string]$Uscita, [string]$Icona, [string[]]$Risorse)
    $provider = New-Object Microsoft.CSharp.CSharpCodeProvider
    $parametri = New-Object System.CodeDom.Compiler.CompilerParameters
    $parametri.GenerateExecutable = $true
    $parametri.OutputAssembly = $Uscita
    $parametri.CompilerOptions = ('/target:winexe /optimize+ /platform:anycpu /win32icon:"{0}"' -f $Icona)
    [void]$parametri.ReferencedAssemblies.AddRange([string[]]@('System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', [psobject].Assembly.Location))
    foreach ($r in $Risorse) { [void]$parametri.EmbeddedResources.Add($r) }
    $esito = $provider.CompileAssemblyFromSource($parametri, $Codice)
    $errori = @($esito.Errors | Where-Object { -not $_.IsWarning })
    if ($errori.Count -gt 0) { throw ('Compilazione non riuscita: ' + (($errori | ForEach-Object { 'riga {0}: {1}' -f $_.Line, $_.ErrorText }) -join '; ')) }
}

function Invoke-CompilaConRoslyn {
    param([string]$Codice, [string]$Uscita, [string]$Icona, [hashtable]$Risorse)
    foreach ($dll in 'Microsoft.CodeAnalysis.dll', 'Microsoft.CodeAnalysis.CSharp.dll') { Add-Type -Path (Join-Path $PSHOME $dll) }
    if (-not ('HDFornitoreRisorsa' -as [type])) {
        Add-Type -TypeDefinition 'public static class HDFornitoreRisorsa { public static System.Func<System.IO.Stream> Da(byte[] b) { return () => new System.IO.MemoryStream(b, false); } }'
    }
    $opzioniTesto = [Microsoft.CodeAnalysis.CSharp.CSharpParseOptions]::new([Microsoft.CodeAnalysis.CSharp.LanguageVersion]::CSharp5)
    $albero = [Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree]::ParseText($Codice, $opzioniTesto)
    $riferimentiDll = foreach ($f in (Get-RiferimentiNet48)) { [Microsoft.CodeAnalysis.MetadataReference]::CreateFromFile($f) }
    $opzioni = [Microsoft.CodeAnalysis.CSharp.CSharpCompilationOptions]::new([Microsoft.CodeAnalysis.OutputKind]::WindowsApplication)
    $opzioni = $opzioni.WithOptimizationLevel([Microsoft.CodeAnalysis.OptimizationLevel]::Release).WithPlatform([Microsoft.CodeAnalysis.Platform]::AnyCpu)
    $compilazione = [Microsoft.CodeAnalysis.CSharp.CSharpCompilation]::Create('ToolkitHelpDesk', [Microsoft.CodeAnalysis.SyntaxTree[]]@($albero), [Microsoft.CodeAnalysis.MetadataReference[]]@($riferimentiDll), $opzioni)
    $descrizioni = foreach ($k in $Risorse.Keys) { [Microsoft.CodeAnalysis.ResourceDescription]::new($k, [HDFornitoreRisorsa]::Da($Risorse[$k]), $true) }
    $flussoIcona = [System.IO.File]::OpenRead($Icona)
    $pe = [System.IO.File]::Create($Uscita)
    try {
        $win32 = $compilazione.CreateDefaultWin32Resources($true, $false, $null, $flussoIcona)
        $esito = $compilazione.Emit($pe, $null, $null, $win32, [Microsoft.CodeAnalysis.ResourceDescription[]]@($descrizioni))
    }
    finally { $pe.Dispose(); $flussoIcona.Dispose() }
    if (-not $esito.Success) {
        Remove-Item -LiteralPath $Uscita -Force -ErrorAction SilentlyContinue
        $errori = @($esito.Diagnostics | Where-Object { $_.Severity -eq [Microsoft.CodeAnalysis.DiagnosticSeverity]::Error })
        throw ('Compilazione non riuscita: ' + (($errori | ForEach-Object { $_.ToString() }) -join '; '))
    }
}

$pacchetto = Get-PacchettoToolkit
if ($SoloPacchetto) { return $pacchetto }

if (-not $Destinazione) { $Destinazione = Join-Path $radice 'ToolkitHelpDesk.exe' }
$Destinazione = [System.IO.Path]::GetFullPath($Destinazione)
$versione = $pacchetto.Versione
if (($versione.Split('.')).Count -lt 4) { $versione = $versione + ('.0' * (4 - ($versione.Split('.')).Count)) }
$codice = (Read-TestoUtf8 (Join-Path $cartellaExe 'Avvio.cs')).Replace('@@VERSIONE@@', $versione).Replace('@@AUTORE@@', $pacchetto.Autore.Replace('"', ''))
$icona = Join-Path $cartellaExe 'icona.ico'
$risorse = [ordered]@{
    'modulo.ps1.gz'      = (ConvertTo-Gzip $pacchetto.Modulo)
    'interfaccia.ps1.gz' = (ConvertTo-Gzip $pacchetto.Interfaccia)
}
$fileProduttori = Join-Path (Join-Path $cartellaModulo 'Dati') 'produttori-mac.txt.gz'
if (Test-Path -LiteralPath $fileProduttori) { $risorse['produttori-mac.txt.gz'] = [System.IO.File]::ReadAllBytes($fileProduttori) }

$temporanea = Join-Path ([System.IO.Path]::GetTempPath()) ('HelpDeskToolkit-exe-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($temporanea)
try {
    $uscitaTemporanea = Join-Path $temporanea 'ToolkitHelpDesk.exe'
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Invoke-CompilaConRoslyn -Codice $codice -Uscita $uscitaTemporanea -Icona $icona -Risorse $risorse
    }
    else {
        $fileRisorse = foreach ($k in $risorse.Keys) {
            $f = Join-Path $temporanea $k
            [System.IO.File]::WriteAllBytes($f, $risorse[$k])
            $f
        }
        Invoke-CompilaConCodeDom -Codice $codice -Uscita $uscitaTemporanea -Icona $icona -Risorse @($fileRisorse)
    }
    Copy-Item -LiteralPath $uscitaTemporanea -Destination $Destinazione -Force
}
finally { Remove-Item -LiteralPath $temporanea -Recurse -Force -ErrorAction SilentlyContinue }

$dimensione = (Get-Item -LiteralPath $Destinazione).Length
Write-Host ('Creato {0} (versione {1}, {2} KB)' -f $Destinazione, $pacchetto.Versione, [math]::Round($dimensione / 1KB)) -ForegroundColor Green
Get-Item -LiteralPath $Destinazione

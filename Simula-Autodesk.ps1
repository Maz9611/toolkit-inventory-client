#Requires -Version 5.1
<#
.SYNOPSIS
SOLO PER UNA VM O PER WINDOWS SANDBOX: simula un PC con prodotti Autodesk per provare il foglio Autodesk.

.DESCRIPTION
Crea dati finti come quelli che lascia un'installazione di Autodesk, così si provano tutti i casi del
foglio Autodesk senza installare niente:
  - AutoCAD 2024 con licenza "utente con nome" e account Autodesk;
  - Civil 3D 2024 con licenza di rete (server licsrv01);
  - Revit 2024 con la licenza di un'altra offerta (una Collection), non avviato da 20 mesi;
  - AutoCAD LT 2019 con il seriale generico delle copie non originali.
Scrive voci in "Installazione applicazioni", chiavi di registro di Autodesk, un finto
AdskLicensingInstHelper.exe (risponde solo a "list"), LoginState.xml dell'utente e due file di Prefetch.

NON usarlo su un PC vero: modifica registro e cartelle di sistema. Con -Rimuovi cancella quello che ha
creato. Si rifiuta di partire se sul PC c'è già Autodesk vero.

Va avviato da Windows PowerShell (powershell.exe) come amministratore:
    powershell -ExecutionPolicy Bypass -File .\Simula-Autodesk.ps1
Poi: Avvia-Prova.cmd (prima rispondendo No alla domanda sull'amministratore, poi Sì).

.PARAMETER Rimuovi
Cancella i dati finti creati da questo script.
#>
[CmdletBinding(SupportsShouldProcess)]
param([switch]$Rimuovi)

$ErrorActionPreference = 'Stop'
$amministratore = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $amministratore) { throw 'Avvialo come amministratore (tasto destro su PowerShell > Esegui come amministratore).' }
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Avvialo con Windows PowerShell (powershell.exe), non con PowerShell 7: serve per creare il finto AdskLicensingInstHelper.exe.' }

$disinstalla = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
$autodesk = 'HKLM:\SOFTWARE\Autodesk'
$cartellaHelper = Join-Path ${env:ProgramFiles(x86)} 'Common Files\Autodesk Shared\AdskLicensing\Current\helper'
$helper = Join-Path $cartellaHelper 'AdskLicensingInstHelper.exe'
$segnoHelper = Join-Path $cartellaHelper 'HDSimulazione.txt'
$loginState = Join-Path $env:LOCALAPPDATA 'Autodesk\Web Services\LoginState.xml'
$prefetch = Join-Path $env:windir 'Prefetch'
$fileAvvii = @('ACAD.EXE-HD000001.pf', 'REVIT.EXE-HD000002.pf')
$prodotti = @(
    @{ Chiave = 'HDSimula-AutoCAD2024'; DisplayName = 'AutoCAD 2024 - English'; DisplayVersion = '24.3.61.0'; Publisher = 'Autodesk, Inc.'; InstallDate = '20240315' }
    @{ Chiave = 'HDSimula-Civil3D2024'; DisplayName = 'Autodesk Civil 3D 2024 - English'; DisplayVersion = '13.6.1580.0'; Publisher = 'Autodesk'; InstallDate = '20240402' }
    @{ Chiave = 'HDSimula-Revit2024'; DisplayName = 'Autodesk Revit 2024'; DisplayVersion = '24.2.0.63'; Publisher = 'Autodesk'; InstallDate = '20240110' }
    @{ Chiave = 'HDSimula-AutoCADLT2019'; DisplayName = 'AutoCAD LT 2019 - Italiano'; DisplayVersion = '23.0.47.0'; Publisher = 'Autodesk'; InstallDate = '20180620' }
    @{ Chiave = 'HDSimula-MaterialLibrary'; DisplayName = 'Autodesk Material Library 2024'; DisplayVersion = '24.0.0'; Publisher = 'Autodesk' }
    @{ Chiave = 'HDSimula-Genuine'; DisplayName = 'Autodesk Genuine Service'; DisplayVersion = '6.0.0'; Publisher = 'Autodesk' }
)

if ($Rimuovi) {
    if (-not $PSCmdlet.ShouldProcess('questo PC', 'Rimuovi i dati Autodesk simulati')) { return }
    foreach ($p in $prodotti) { Remove-Item -LiteralPath (Join-Path $disinstalla $p.Chiave) -Recurse -Force -ErrorAction SilentlyContinue }
    if ((Get-ItemProperty -LiteralPath $autodesk -Name HDSimulazione -ErrorAction SilentlyContinue).HDSimulazione -eq 1) { Remove-Item -LiteralPath $autodesk -Recurse -Force }
    if (Test-Path -LiteralPath $segnoHelper) { Remove-Item -LiteralPath (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $cartellaHelper))) -Recurse -Force }
    if ((Test-Path -LiteralPath $loginState) -and (Get-Content -LiteralPath $loginState -Raw) -match 'HDSimulazione') { Remove-Item -LiteralPath $loginState -Force }
    foreach ($f in $fileAvvii) { Remove-Item -LiteralPath (Join-Path $prefetch $f) -Force -ErrorAction SilentlyContinue }
    Write-Host 'Dati Autodesk simulati rimossi.' -ForegroundColor Green
    return
}

# Mai sopra un'installazione vera di Autodesk
$giaSimulato = (Get-ItemProperty -LiteralPath $autodesk -Name HDSimulazione -ErrorAction SilentlyContinue).HDSimulazione -eq 1
if (((Test-Path -LiteralPath $autodesk) -and -not $giaSimulato) -or ((Test-Path -LiteralPath $helper) -and -not (Test-Path -LiteralPath $segnoHelper))) {
    throw 'Su questo PC c''è già Autodesk vero: la simulazione serve su un PC (o una VM) senza Autodesk.'
}
if (-not $PSCmdlet.ShouldProcess('questo PC', 'Crea i dati Autodesk simulati (registro, Program Files, Prefetch)')) { return }

# 1. "Installazione applicazioni"
foreach ($p in $prodotti) {
    $k = New-Item -Path $disinstalla -Name $p.Chiave -Force
    foreach ($nome in 'DisplayName', 'DisplayVersion', 'Publisher', 'InstallDate') {
        if ($p[$nome]) { [void](New-ItemProperty -LiteralPath $k.PSPath -Name $nome -Value $p[$nome] -PropertyType String -Force) }
    }
}

# 2. Chiave di AutoCAD LT 2019 con il seriale (le versioni a seriale lo scrivono qui)
[void](New-Item -Path $autodesk -Force)
[void](New-ItemProperty -LiteralPath $autodesk -Name HDSimulazione -Value 1 -PropertyType DWord -Force)
$lt = New-Item -Path (Join-Path $autodesk 'AutoCAD LT\R23.0\ACADLT-3001:410') -Force
[void](New-ItemProperty -LiteralPath $lt.PSPath -Name ProductName -Value 'AutoCAD LT 2019' -PropertyType String -Force)
[void](New-ItemProperty -LiteralPath $lt.PSPath -Name SerialNumber -Value '666-69696969' -PropertyType String -Force)

# 3. Finto AdskLicensingInstHelper.exe: risponde a "list" come quello vero
[void](New-Item -ItemType Directory -Path $cartellaHelper -Force)
Set-Content -LiteralPath $segnoHelper -Value 'Creato da Simula-Autodesk.ps1 (Toolkit Help Desk): cancellalo con Simula-Autodesk.ps1 -Rimuovi' -Encoding UTF8
$codice = @'
public static class HelperSimulato {
    public static int Main(string[] args) {
        if (args.Length > 0 && args[0] == "list") {
            System.Console.WriteLine(@"[
  { ""feature_id"": ""ACD"", ""def_prod_key"": ""001P1"", ""def_prod_ver"": ""2024.0.0.F"", ""sel_prod_key"": ""001P1"", ""sel_prod_ver"": ""2024.0.0.F"", ""sel_mode"": ""USER"", ""sel_srv_names"": [] },
  { ""feature_id"": ""CIV3D"", ""def_prod_key"": ""237P1"", ""def_prod_ver"": ""2024.0.0.F"", ""sel_prod_key"": ""237P1"", ""sel_prod_ver"": ""2024.0.0.F"", ""sel_mode"": ""NETWORK"", ""sel_srv_names"": [""licsrv01""] },
  { ""feature_id"": ""RVT"", ""def_prod_key"": ""829P1"", ""def_prod_ver"": ""2024.0.0.F"", ""sel_prod_key"": ""02HP1"", ""sel_prod_ver"": ""2024.0.0.F"", ""sel_mode"": ""USER"", ""sel_srv_names"": null }
]");
            return 0;
        }
        System.Console.Error.WriteLine("Helper simulato: conosce solo il comando list");
        return 1;
    }
}
'@
if (Test-Path -LiteralPath $helper) { Remove-Item -LiteralPath $helper -Force }
Add-Type -TypeDefinition $codice -OutputAssembly $helper -OutputType ConsoleApplication

# 4. Account Autodesk dell'utente (LoginState.xml)
[void](New-Item -ItemType Directory -Path (Split-Path -Parent $loginState) -Force)
$xml = '<?xml version="1.0" encoding="utf-8"?><!-- HDSimulazione --><LoginState><UserInfo><UserName>mrossi</UserName><Email>mario.rossi@example.com</Email></UserInfo></LoginState>'
[System.IO.File]::WriteAllText($loginState, $xml, (New-Object System.Text.UTF8Encoding($false)))

# 5. Ultimo avvio: AutoCAD (e Civil 3D, stesso programma) 40 giorni fa, Revit 20 mesi fa
$avvioAcad = (Get-Date).AddDays(-40)
$avvioRevit = (Get-Date).AddMonths(-20)
if (Test-Path -LiteralPath $prefetch) {
    foreach ($coppia in @(@($fileAvvii[0], $avvioAcad), @($fileAvvii[1], $avvioRevit))) {
        $f = Join-Path $prefetch $coppia[0]
        [System.IO.File]::WriteAllBytes($f, [byte[]]@(0))
        (Get-Item -LiteralPath $f).LastWriteTime = $coppia[1]
    }
}
else { Write-Warning 'Cartella Prefetch assente: la colonna "Ultimo avvio" resterà vuota.' }

$data = { param($d) $d.ToString('dd/MM/yyyy') }
Write-Host ''
Write-Host 'Dati Autodesk simulati creati. Ora lancia Avvia-Prova.cmd e apri l''Excel di prova.' -ForegroundColor Green
Write-Host ''
Write-Host 'Da AMMINISTRATORE nel foglio Autodesk devi vedere 4 righe:' -ForegroundColor Cyan
Write-Host ('  AutoCAD 2024     Utente con nome (abbonamento, account Autodesk); mario.rossi@example.com; ultimo avvio {0} (uno dei 2 prodotti con lo stesso programma)' -f (& $data $avvioAcad))
Write-Host ('  AutoCAD LT 2019  Singola con seriale (versione 2019 o precedente); seriale 666-69696969 (seriale delle copie non originali); ultimo avvio: Mai')
Write-Host ('  Civil 3D 2024    Di rete (server licsrv01); ultimo avvio {0} (uno dei 2 prodotti con lo stesso programma)' -f (& $data $avvioAcad))
Write-Host ('  Revit 2024       Utente con nome (...), con la licenza di un''altra offerta (chiave 02HP1, es. una Collection); ultimo avvio {0}' -f (& $data $avvioRevit))
Write-Host 'Nel foglio Riepilogo, nella riga di questo PC (insieme ai controlli su Windows):' -ForegroundColor Cyan
Write-Host '  Urgente:        AutoCAD LT 2019: seriale delle copie non originali (666-69696969)'
Write-Host ('  Da controllare: Revit 2024 non avviato dal {0}' -f (& $data $avvioRevit))
Write-Host 'Da UTENTE NORMALE: stesse righe, ma "Ultimo avvio" = "Serve avviare come amministratore".' -ForegroundColor Cyan
Write-Host 'Niente "Material Library" né "Genuine Service": sono componenti, non prodotti.' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Per togliere tutto: powershell -ExecutionPolicy Bypass -File .\Simula-Autodesk.ps1 -Rimuovi' -ForegroundColor DarkGray

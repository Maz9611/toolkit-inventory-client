# Spazio disco: volumi, cartelle più grandi, stima dello spazio recuperabile.
# Tutto in SOLA LETTURA: nessuna funzione cancella file.

function Get-HDStatoSpazio {
    # OK / ATTENZIONE / CRITICO in base a GB liberi e percentuale libera
    param(
        [Parameter(Mandatory)][double]$LiberoGB,
        [Parameter(Mandatory)][double]$LiberoPct,
        [hashtable]$Soglie
    )
    if (-not $Soglie) { $Soglie = (Get-HDConfig).Soglie }
    if ($LiberoGB -lt $Soglie.DiscoCriticoGB -or $LiberoPct -lt $Soglie.DiscoCriticoPct) { return 'CRITICO' }
    if ($LiberoGB -lt $Soglie.DiscoAttenzioneGB -or $LiberoPct -lt $Soglie.DiscoAttenzionePct) { return 'ATTENZIONE' }
    'OK'
}

function ConvertTo-HDVolume {
    # Win32_LogicalDisk -> oggetto leggibile
    param([Parameter(Mandatory)]$Disco, [string]$Computer, [hashtable]$Soglie)
    if (-not $Computer) { $Computer = Get-HDNomePC }
    $dimensione = [double]$Disco.Size
    $libero = [double]$Disco.FreeSpace
    $pct = 0
    if ($dimensione -gt 0) { $pct = [math]::Round($libero / $dimensione * 100, 1) }
    $liberoGB = [math]::Round($libero / 1GB, 1)
    [pscustomobject]@{
        PSTypeName = 'HelpDesk.Volume'
        Computer   = $Computer
        Unita      = $Disco.DeviceID
        Etichetta  = $Disco.VolumeName
        FileSystem = $Disco.FileSystem
        TotaleGB   = [math]::Round($dimensione / 1GB, 1)
        UsatoGB    = [math]::Round(($dimensione - $libero) / 1GB, 1)
        LiberoGB   = $liberoGB
        LiberoPct  = $pct
        Stato      = Get-HDStatoSpazio -LiberoGB $liberoGB -LiberoPct $pct -Soglie $Soglie
    }
}

function Get-SpazioDisco {
    <#
    .SYNOPSIS
    Spazio libero dei dischi, con stato OK / ATTENZIONE / CRITICO.

    .DESCRIPTION
    Legge i volumi locali (o di PC remoti via CIM/WinRM) e valuta lo spazio libero
    con le soglie di config.psd1. Non modifica nulla.

    .PARAMETER ComputerName
    Uno o più PC remoti (serve WinRM attivo e un account amministratore).

    .PARAMETER IncludiRimovibili
    Considera anche chiavette e dischi USB.

    .EXAMPLE
    Get-SpazioDisco

    .EXAMPLE
    Get-SpazioDisco | Where-Object Stato -ne 'OK'

    .EXAMPLE
    Get-SpazioDisco -ComputerName PC-01, PC-02
    #>
    [CmdletBinding()]
    param(
        [string[]]$ComputerName,
        [switch]$IncludiRimovibili
    )
    $soglie = (Get-HDConfig).Soglie
    $filtro = 'DriveType=3'
    if ($IncludiRimovibili) { $filtro = 'DriveType=2 OR DriveType=3' }
    $destinazioni = @('')
    if ($ComputerName) { $destinazioni = $ComputerName }

    foreach ($pc in $destinazioni) {
        try {
            $dischi = @(Get-HDCim -Classe 'Win32_LogicalDisk' -Filtro $filtro -ComputerName $pc)
        }
        catch {
            $nome = $pc
            if (-not $nome) { $nome = Get-HDNomePC }
            Write-Warning "Impossibile leggere i dischi di '$nome': $($_.Exception.Message)"
            continue
        }
        $nomePC = $pc
        if (-not $nomePC) { $nomePC = Get-HDNomePC }
        foreach ($d in $dischi) {
            if ([double]$d.Size -gt 0) { ConvertTo-HDVolume -Disco $d -Computer $nomePC -Soglie $soglie }
        }
    }
}

function Measure-HDCartella {
    # Somma la dimensione dei file di una cartella e sottocartelle.
    # Non segue i collegamenti (junction/symlink) e conta le cartelle non leggibili.
    param([Parameter(Mandatory)][string]$Percorso)
    $byte = [double]0
    $file = 0
    $errori = 0
    $pila = New-Object 'System.Collections.Generic.Stack[string]'
    $pila.Push($Percorso)
    while ($pila.Count -gt 0) {
        $dir = $pila.Pop()
        try {
            foreach ($voce in (New-Object System.IO.DirectoryInfo($dir)).EnumerateFileSystemInfos()) {
                if ($voce -is [System.IO.DirectoryInfo]) {
                    if (-not ($voce.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) { $pila.Push($voce.FullName) }
                }
                else {
                    $byte += $voce.Length
                    $file++
                }
            }
        }
        catch { $errori++ }
    }
    [pscustomobject]@{ Byte = $byte; File = $file; Errori = $errori }
}

function Get-CartelleGrandi {
    <#
    .SYNOPSIS
    Elenca le sottocartelle più grandi di un percorso (per capire dove è finito lo spazio).

    .PARAMETER Percorso
    Cartella da analizzare. Predefinito: la cartella del profilo utente.

    .PARAMETER Primi
    Quante cartelle mostrare (predefinito 10).

    .EXAMPLE
    Get-CartelleGrandi C:\Users

    .EXAMPLE
    Get-CartelleGrandi -Percorso D:\Dati -Primi 20
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][string]$Percorso,
        [ValidateRange(1, 200)][int]$Primi = 10
    )
    if (-not $Percorso) {
        $Percorso = $env:USERPROFILE
        if (-not $Percorso) { $Percorso = $HOME }
    }
    $base = Resolve-HDPercorso $Percorso
    if (-not (Test-Path -LiteralPath $base -PathType Container)) { throw "Cartella non trovata: $Percorso" }

    $sotto = @(Get-ChildItem -LiteralPath $base -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) })
    $risultati = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($c in $sotto) {
        $i++
        Write-Progress -Activity "Dimensione cartelle in $base" -Status $c.Name -PercentComplete ([int]($i * 100 / $sotto.Count))
        $m = Measure-HDCartella -Percorso $c.FullName
        $risultati.Add([pscustomobject]@{
                PSTypeName    = 'HelpDesk.Cartella'
                Cartella      = $c.FullName
                DimensioneGB  = [math]::Round($m.Byte / 1GB, 2)
                Dimensione    = Format-HDByte $m.Byte
                File          = $m.File
                NonLeggibili  = $m.Errori
                Byte          = $m.Byte
            })
    }
    Write-Progress -Activity "Dimensione cartelle in $base" -Completed

    # i file sparsi direttamente nella cartella analizzata
    $sciolti = @(Get-ChildItem -LiteralPath $base -File -Force -ErrorAction SilentlyContinue)
    if ($sciolti.Count -gt 0) {
        $somma = [double](($sciolti | Measure-Object -Property Length -Sum).Sum)
        $risultati.Add([pscustomobject]@{
                PSTypeName    = 'HelpDesk.Cartella'
                Cartella      = (Join-Path $base '(file nella cartella)')
                DimensioneGB  = [math]::Round($somma / 1GB, 2)
                Dimensione    = Format-HDByte $somma
                File          = $sciolti.Count
                NonLeggibili  = 0
                Byte          = $somma
            })
    }
    $risultati | Sort-Object Byte -Descending | Select-Object -First $Primi
}

function Get-HDVociPulizia {
    # Posizioni tipiche dove si recupera spazio. Separate per poterle simulare nei test.
    $voci = New-Object System.Collections.Generic.List[object]
    $win = $env:windir
    $sys = $env:SystemDrive
    if ($env:TEMP) {
        $voci.Add(@{ Voce = 'File temporanei utente'; Percorso = $env:TEMP; Come = 'Impostazioni > Sistema > Archiviazione > File temporanei' })
    }
    if ($win) {
        $voci.Add(@{ Voce = 'File temporanei di Windows'; Percorso = (Join-Path $win 'Temp'); Come = 'Pulizia disco (cleanmgr) avviata come amministratore' })
        $voci.Add(@{ Voce = 'Cache download di Windows Update'; Percorso = (Join-Path $win 'SoftwareDistribution\Download'); Come = 'Pulizia disco > Pulizia file di sistema > Pulizia di Windows Update' })
        $voci.Add(@{ Voce = 'Dump di memoria (minidump)'; Percorso = (Join-Path $win 'Minidump'); Come = 'Pulizia disco > File di dump della memoria (prima analizzali se ci sono BSOD)' })
        $voci.Add(@{ Voce = 'Dump di memoria completo'; Percorso = (Join-Path $win 'MEMORY.DMP'); Come = 'Pulizia disco > File di dump della memoria' })
    }
    if ($env:ProgramData) {
        $voci.Add(@{ Voce = 'Segnalazioni errori Windows'; Percorso = (Join-Path $env:ProgramData 'Microsoft\Windows\WER'); Come = 'Pulizia disco > File di segnalazione errori di Windows' })
    }
    if ($sys) {
        $voci.Add(@{ Voce = 'Installazione precedente (Windows.old)'; Percorso = (Join-Path $sys 'Windows.old'); Come = 'Pulizia disco > Installazioni precedenti di Windows (dopo, niente più ritorno alla versione precedente)' })
        $voci.Add(@{ Voce = 'Cestino (tutti gli utenti)'; Percorso = (Join-Path $sys '$Recycle.Bin'); Come = 'Svuota il Cestino, dopo averlo chiesto all''utente' })
    }
    if ($env:USERPROFILE) {
        $voci.Add(@{ Voce = 'Cartella Download dell''utente'; Percorso = (Join-Path $env:USERPROFILE 'Downloads'); Come = 'Solo da valutare CON l''utente: non cancellare senza permesso' })
    }
    $voci
}

function Get-StimaPulizia {
    <#
    .SYNOPSIS
    Stima quanto spazio si può recuperare (Temp, Windows Update, Cestino, Windows.old...).

    .DESCRIPTION
    Misura soltanto: non cancella nulla. Per alcune cartelle servono i diritti di
    amministratore; in quel caso la colonna NonLeggibili è maggiore di zero e il
    valore è sottostimato.

    .EXAMPLE
    Get-StimaPulizia | Format-Table Voce, Dimensione, Come -AutoSize
    #>
    [CmdletBinding()]
    param()
    foreach ($v in (Get-HDVociPulizia)) {
        $p = $v.Percorso
        $byte = [double]0
        $errori = 0
        $trovato = $true
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $byte = [double](Get-Item -LiteralPath $p -Force).Length
        }
        elseif (Test-Path -LiteralPath $p -PathType Container) {
            $m = Measure-HDCartella -Percorso $p
            $byte = $m.Byte
            $errori = $m.Errori
        }
        else { $trovato = $false }
        if (-not $trovato) { continue }
        $nota = ''
        if ($errori -gt 0) { $nota = "$errori cartelle non leggibili: avvia come amministratore per il valore completo" }
        [pscustomobject]@{
            PSTypeName   = 'HelpDesk.Pulizia'
            Voce         = $v.Voce
            Dimensione   = Format-HDByte $byte
            DimensioneGB = [math]::Round($byte / 1GB, 2)
            Come         = $v.Come
            Percorso     = $p
            NonLeggibili = $errori
            Nota         = $nota
            Byte         = $byte
        }
    }
}

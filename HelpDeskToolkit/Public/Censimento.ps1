# Censimento PC su Excel: import della scansione di rete e completamento PC per PC.

# Raccolta dei dati grezzi sul PC. Blocco autonomo (niente funzioni del modulo) così
# gira anche su un PC remoto con Invoke-Command. Sola lettura.
$script:HDRaccoltaCensimento = {
    param([string]$ModelloCynet = 'Cynet')
    function Leggi([scriptblock]$Blocco) { try { & $Blocco } catch { $null } }
    $cronometro = [System.Diagnostics.Stopwatch]::StartNew()
    $g = [ordered]@{}

    # Le letture lente partono subito in parallelo (runspace separati, testo autonomo) mentre
    # il resto si legge qui; se il parallelo non è disponibile si fanno in fila alla fine.
    $lente = [ordered]@{
        Licenza   = {
            $t = [System.Diagnostics.Stopwatch]::StartNew()
            $r = $null
            try { $r = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction Stop | Select-Object Name, Description, LicenseStatus) } catch { $r = $null }
            @{ Licenza = $r; Ms = $t.ElapsedMilliseconds }
        }
        Programmi = {
            $t = [System.Diagnostics.Stopwatch]::StartNew()
            $chiavi = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            $p = $null
            try { $p = @(Get-ItemProperty -Path $chiavi -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, SystemComponent, ParentKeyName, ReleaseType) } catch { $p = $null }
            $o = $null
            try { $o = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop | Select-Object ProductReleaseIds, VersionToReport } catch { $o = $null }
            @{ Programmi = $p; OfficeC2R = $o; Ms = $t.ElapsedMilliseconds }
        }
        Dischi    = {
            $t = [System.Diagnostics.Stopwatch]::StartNew()
            $r = @{ DischiFisici = $null; Smart = $null; SmartPrevisione = $null; DischiWin32 = $null }
            $fisici = $null
            $amministratore = $false
            try { $amministratore = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $amministratore = $false }
            try { $fisici = @(Get-CimInstance -Namespace 'root/Microsoft/Windows/Storage' -ClassName MSFT_PhysicalDisk -ErrorAction Stop) } catch { $fisici = $null }
            if ($null -ne $fisici) {
                $r.DischiFisici = @($fisici | Select-Object FriendlyName, MediaType, BusType, Size, HealthStatus)
                # SMART: contatori di affidabilità (temperatura, usura, ore, errori). Servono i diritti di
                # amministratore: senza, Windows restituisce contatori vuoti, quindi non si leggono.
                if ($amministratore) {
                    try {
                        $r.Smart = @($fisici | Where-Object { [string]$_.BusType -ne '7' } | ForEach-Object {
                                $c = Get-CimAssociatedInstance -InputObject $_ -ResultClassName MSFT_StorageReliabilityCounter -ErrorAction Stop | Select-Object -First 1
                                [pscustomobject]@{ Disco = [string]$_.FriendlyName; Temperatura = $c.Temperature; Usura = $c.Wear; Ore = $c.PowerOnHours; ErroriLettura = $c.ReadErrorsUncorrected; ErroriScrittura = $c.WriteErrorsUncorrected }
                            })
                    }
                    catch { $r.Smart = $null }
                }
            }
            else {
                try { $r.DischiWin32 = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop | Select-Object Model, Size, MediaType, InterfaceType, Status) } catch { $r.DischiWin32 = $null }
            }
            try { $r.SmartPrevisione = @(Get-CimInstance -Namespace 'root/wmi' -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction Stop | Select-Object InstanceName, PredictFailure) } catch { $r.SmartPrevisione = $null }
            $r.Ms = $t.ElapsedMilliseconds
            $r
        }
    }
    $inCorso = @{}
    $pool = $null
    try {
        $pool = [runspacefactory]::CreateRunspacePool(1, $lente.Count)
        $pool.Open()
        foreach ($k in $lente.Keys) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($lente[$k].ToString())
            $inCorso[$k] = @{ PS = $ps; Attesa = $ps.BeginInvoke() }
        }
    }
    catch { $inCorso = @{} }

    $g.NomePC = $env:COMPUTERNAME
    $g.UnitaSistema = $env:SystemDrive
    $g.ComputerSystem = Leggi { Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop | Select-Object -First 1 Manufacturer, Model, Domain, PartOfDomain, Workgroup, UserName, TotalPhysicalMemory, PCSystemType }
    $g.Prodotto = Leggi { Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop | Select-Object -First 1 Name, Version, IdentifyingNumber }
    $g.SistemaOperativo = Leggi { Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop | Select-Object -First 1 Caption, Version, BuildNumber }
    $g.VersioneWindows = Leggi { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop | Select-Object DisplayVersion, ReleaseId, CurrentBuild, UBR }
    $g.Bios = Leggi { Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop | Select-Object -First 1 SerialNumber, SMBIOSBIOSVersion, ReleaseDate, Manufacturer }
    $g.Telaio = Leggi { Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1 SerialNumber }
    $g.SchedaMadre = Leggi { Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction Stop | Select-Object -First 1 Manufacturer, Product, SerialNumber }
    $g.Chassis = Leggi { @((Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1).ChassisTypes) }
    $g.Processori = Leggi { @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object Name) }
    $g.Memoria = Leggi { @(Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop | Select-Object Capacity, FormFactor, SMBIOSMemoryType, MemoryType, Speed) }
    $g.Volumi = Leggi { @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop | Select-Object DeviceID, Size, FreeSpace) }
    $g.ConfigRete = Leggi { @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop | Select-Object Description, IPAddress, MACAddress, DefaultIPGateway, DHCPEnabled, IPConnectionMetric) }
    $g.SchedeRete = Leggi { @(Get-CimInstance -ClassName Win32_NetworkAdapter -Filter 'PhysicalAdapter=True' -ErrorAction Stop | Where-Object { $_.MACAddress } | Select-Object Name, MACAddress, NetConnectionStatus) }
    $g.UltimoAccesso = Leggi { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI' -ErrorAction Stop).LastLoggedOnSAMUser }
    $g.Profili = Leggi { @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop | Select-Object LocalPath, LastUseTime, Special) }
    $g.Entra = Leggi {
        $adesione = @(Get-ChildItem -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo' -ErrorAction Stop)
        if ($adesione.Count -eq 0) { return $null }
        $tenant = Get-ChildItem -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo' -ErrorAction SilentlyContinue | Select-Object -First 1
        $nome = $null
        if ($tenant) { $nome = (Get-ItemProperty -Path $tenant.PSPath -ErrorAction SilentlyContinue).DisplayName }
        if ($nome) { $nome } else { '-' }
    }
    # Dati per il foglio "Dettagli": pronto per Windows 11, dischi, monitor, stampanti, software.
    # Senza diritti di amministratore alcuni (TPM, BitLocker) si leggono in modo meno preciso.
    $g.Avvio = Leggi {
        if ($env:firmware_type -eq 'UEFI') { 2 }
        elseif ($env:firmware_type -eq 'Legacy') { 1 }
        else { (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name PEFirmwareType -ErrorAction Stop).PEFirmwareType }
    }
    $g.SecureBoot = Leggi { [pscustomobject]@{ Attivo = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' -Name UEFISecureBootEnabled -ErrorAction Stop).UEFISecureBootEnabled } }
    $g.Tpm = Leggi { Get-CimInstance -Namespace 'root/cimv2/Security/MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop | Select-Object -First 1 SpecVersion, IsEnabled_InitialValue, IsActivated_InitialValue }
    $g.TpmDispositivi = Leggi { [pscustomobject]@{ Nomi = @(Get-CimInstance -ClassName Win32_PnPEntity -Filter "PNPClass='SecurityDevices'" -ErrorAction Stop | ForEach-Object { [string]$_.Name }) } }
    $g.BitLockerAdmin = Leggi { @(Get-CimInstance -Namespace 'root/cimv2/Security/MicrosoftVolumeEncryption' -ClassName Win32_EncryptableVolume -ErrorAction Stop | Select-Object DriveLetter, ProtectionStatus) }
    if ($null -eq $g.BitLockerAdmin) {
        $g.BitLockerShell = Leggi {
            $shell = New-Object -ComObject Shell.Application
            @(@($g.Volumi) | Where-Object { $_.DeviceID } | ForEach-Object {
                    [pscustomobject]@{ Unita = [string]$_.DeviceID; Valore = $shell.NameSpace([string]$_.DeviceID + '\').Self.ExtendedProperty('System.Volume.BitLockerProtection') }
                })
        }
    }
    $g.Monitor = Leggi {
        $testo = { param($Codici) -join @(@($Codici) | Where-Object { $_ -gt 0 } | ForEach-Object { [char][int]$_ }) }
        @(Get-CimInstance -Namespace 'root/wmi' -ClassName WmiMonitorID -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ Produttore = (& $testo $_.ManufacturerName); Nome = (& $testo $_.UserFriendlyName); Seriale = (& $testo $_.SerialNumberID); Anno = $_.YearOfManufacture }
            })
    }
    $g.Stampanti = Leggi { @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Select-Object Name, DriverName, PortName, Network, Default) }
    $g.Antivirus = Leggi { @(Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop | Select-Object displayName, productState) }
    $g.DesktopRemoto = Leggi { (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -ErrorAction Stop).fDenyTSConnections }
    $percorsiCynet = @()
    if ($env:windir) { $percorsiCynet += Join-Path $env:windir 'CynetEPS.exe' }
    if ($env:ProgramFiles) { $percorsiCynet += Join-Path $env:ProgramFiles 'Cynet' }
    if (${env:ProgramFiles(x86)}) { $percorsiCynet += Join-Path ${env:ProgramFiles(x86)} 'Cynet' }
    $g.Cynet = [pscustomobject]@{
        Processi  = @(Leggi { Get-Process -ErrorAction Stop | Where-Object { $_.ProcessName -match $ModelloCynet } | ForEach-Object { $_.ProcessName } | Select-Object -Unique })
        Servizi   = @(Leggi { Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -match $ModelloCynet -or $_.DisplayName -match $ModelloCynet } | ForEach-Object { $_.Name } })
        Programmi = @()
        File      = @($percorsiCynet | Where-Object { Test-Path -LiteralPath $_ })
    }
    $tempi = [ordered]@{ Principale = $cronometro.ElapsedMilliseconds }

    # Risultati delle letture lente
    foreach ($k in $lente.Keys) {
        $risultato = $null
        if ($inCorso.ContainsKey($k)) {
            try { $risultato = @($inCorso[$k].PS.EndInvoke($inCorso[$k].Attesa))[0] } catch { $risultato = $null }
            $inCorso[$k].PS.Dispose()
        }
        else {
            try { $risultato = & $lente[$k] } catch { $risultato = $null }
        }
        if ($risultato -is [hashtable]) {
            foreach ($campo in $risultato.Keys) { if ($campo -ne 'Ms') { $g[$campo] = $risultato[$campo] } }
            $tempi[$k] = $risultato.Ms
        }
    }
    if ($pool) { $pool.Close(); $pool.Dispose() }
    if ($g.Programmi) {
        $g.Cynet.Programmi = @(@($g.Programmi) | Where-Object { $_.DisplayName -match $ModelloCynet -or $_.Publisher -match $ModelloCynet } | ForEach-Object { $_.DisplayName } | Select-Object -Unique)
    }
    $tempi['Totale'] = $cronometro.ElapsedMilliseconds
    $g.Tempi = [pscustomobject]$tempi
    $g.Raccolto = Get-Date
    [pscustomobject]$g
}

function Invoke-HDRaccoltaInventario {
    param([string]$ComputerName, [pscredential]$Credential, [string]$ModelloCynet = 'Cynet')
    $locale = (-not $ComputerName) -or ($ComputerName -in @('.', 'localhost', '127.0.0.1', (Get-HDNomePC)))
    if ($locale) {
        if (-not (Test-HDWindows)) { throw 'Il censimento si fa su Windows.' }
        return (& $script:HDRaccoltaCensimento $ModelloCynet)
    }
    $p = @{ ComputerName = $ComputerName; ScriptBlock = $script:HDRaccoltaCensimento; ArgumentList = @($ModelloCynet); ErrorAction = 'Stop' }
    if ($Credential) { $p.Credential = $Credential }
    Invoke-Command @p
}

function Test-HDInterattivo {
    if (-not [Environment]::UserInteractive) { return $false }
    if ($Host.Name -eq 'ToolkitHelpDesk') { return $false }   # dentro l'exe non c'è una console per le domande
    if ([Environment]::GetCommandLineArgs() -contains '-NonInteractive') { return $false }
    $true
}

function Get-InventarioPC {
    <#
    .SYNOPSIS
    Legge i dati del PC per il censimento, senza scrivere nulla.

    .DESCRIPTION
    Restituisce le colonne del censimento: Hostname, Indirizzo IP, Tipo IP, MAC Address,
    Device Type, Manufacturer, Model, Nome Utente, Dominio, Serial Number, Operating System
    e versione, Memory, Processor Type, Free Disk Space % e Logical Disk Size del disco di
    sistema, Cynet (Sì/No). Più i dettagli usati per ritrovare il PC nell'Excel e, nelle
    proprietà Dettagli e Software, i dati dei fogli Dettagli (Windows 11, TPM, dischi, monitor,
    stampanti, Office, licenza, antivirus, accesso remoto) e Software (programmi installati).

    .PARAMETER ComputerName
    PC remoti (serve PowerShell Remoting/WinRM attivo e un account amministratore).

    .EXAMPLE
    Get-InventarioPC
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline)][string[]]$ComputerName,
        [pscredential]$Credential
    )
    begin { $cfg = Get-HDConfig }
    process {
        $destinazioni = @('')
        if ($ComputerName) { $destinazioni = $ComputerName }
        foreach ($pc in $destinazioni) {
            try {
                $grezzo = Invoke-HDRaccoltaInventario -ComputerName $pc -Credential $Credential -ModelloCynet $cfg.Censimento.RilevaCynet
                ConvertTo-HDRecordCensimento -Grezzo $grezzo -AccountTecnici $cfg.Censimento.AccountTecnici -TipoClient $cfg.Censimento.DeviceTypeClient -TipoServer $cfg.Censimento.DeviceTypeServer -SenzaDominio $cfg.Censimento.DominioAssente
            }
            catch {
                $nome = $pc
                if (-not $nome) { $nome = Get-HDNomePC }
                Write-Warning "Raccolta dati non riuscita per '$nome': $($_.Exception.Message)"
            }
        }
    }
}

#region File del censimento

function Get-HDElencoCensimenti {
    param([string]$Cartella)
    @(Get-ChildItem -LiteralPath $Cartella -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '.~*' -and $_.Name -ne 'Modello.xlsx' } |
            Sort-Object LastWriteTime -Descending)
}

function Get-HDModelloCensimento {
    $cfg = Get-HDConfig
    if ($cfg.Censimento.ModelloExcel) {
        $m = [Environment]::ExpandEnvironmentVariables([string]$cfg.Censimento.ModelloExcel)
        if (Test-Path -LiteralPath $m) { return (Resolve-HDPercorso $m) }
        Write-Warning "Modello Excel indicato in config.psd1 non trovato: $m. Uso il modello standard."
        return $null
    }
    $predefinito = Join-Path (Get-HDCartella -Tipo Censimento) 'Modello.xlsx'
    if (Test-Path -LiteralPath $predefinito) { return $predefinito }
    $null
}

function Get-HDFogliPrincipali {
    # I fogli in cui cercare il censimento: tutti tranne Dettagli, Software e DatiSoftware
    param([object[]]$Fogli)
    $nostri = @($script:HDFoglioDettagli, $script:HDFoglioSoftware, $script:HDFoglioDatiSoftware)
    $principali = @($Fogli | Where-Object { $nostri -notcontains $_.Nome })
    if ($principali.Count -eq 0) { return $Fogli }
    $principali
}

function Get-HDTabellaCensimento {
    # Foglio principale del censimento, più i fogli Dettagli, DatiSoftware e (stato del) Software
    param([Parameter(Mandatory)][string]$Percorso)
    $p = Resolve-HDPercorso $Percorso
    if ([System.IO.Path]::GetExtension($p) -ne '.xlsx') {
        throw "Il censimento deve essere un file .xlsx (in Excel: File > Salva con nome > Cartella di lavoro di Excel): $p"
    }
    $fogli = @(Read-HDXlsx -Percorso $p)
    $t = ConvertTo-HDTabellaCensimento -Fogli (Get-HDFogliPrincipali $fogli) -Percorso $p
    $t | Add-Member -NotePropertyName RigheIniziali -NotePropertyValue $t.Righe.Count
    $t | Add-Member -NotePropertyName Dettagli -NotePropertyValue (ConvertTo-HDTabellaFoglio -Fogli $fogli -Nome $script:HDFoglioDettagli -Colonne $script:HDColonneDettagli)
    $t | Add-Member -NotePropertyName DatiSoftware -NotePropertyValue (ConvertTo-HDTabellaFoglio -Fogli $fogli -Nome $script:HDFoglioDatiSoftware -Colonne $script:HDColonneDatiSoftware -Nascosto)
    $t | Add-Member -NotePropertyName Software -NotePropertyValue (Get-HDStatoFoglioSoftware $fogli)
    $t
}

function Resolve-HDFileCensimento {
    # Quale Excel usare: quello indicato, quello del cliente, l'unico della cartella, oppure si sceglie
    param([string]$Percorso, [string]$Cliente, [switch]$Crea, [switch]$NonChiedere)
    if ($Percorso) {
        $p = Resolve-HDPercorso $Percorso
        if (Test-Path -LiteralPath $p) { return $p }
        if ($Crea) { return (New-Censimento -Percorso $p -Confirm:$false).FullName }
        throw "File del censimento non trovato: $Percorso"
    }
    $cartella = Get-HDCartella -Tipo Censimento
    $file = @(Get-HDElencoCensimenti $cartella)
    if ($Cliente) {
        $delCliente = $file | Where-Object { $_.BaseName -eq ('Censimento_' + (ConvertTo-HDNomeFile $Cliente)) } | Select-Object -First 1
        if ($delCliente) { return $delCliente.FullName }
    }
    elseif ($file.Count -eq 1) { return $file[0].FullName }
    elseif ($file.Count -gt 1) {
        if ($NonChiedere) { return $file[0].FullName }
        Write-Host 'Nella chiavetta ci sono più censimenti:'
        for ($i = 0; $i -lt $file.Count; $i++) {
            Write-Host ('  {0}  {1}   (modificato il {2})' -f ($i + 1), $file[$i].Name, (Format-HDData $file[$i].LastWriteTime))
        }
        $s = Read-Host 'Quale uso? (INVIO = 1)'
        if (-not $s) { return $file[0].FullName }
        if ($s -match '^\d+$' -and [int]$s -ge 1 -and [int]$s -le $file.Count) { return $file[[int]$s - 1].FullName }
        throw 'Scelta non valida.'
    }
    if (-not $Crea) { throw "Nessun censimento in $cartella`: importa prima la scansione di rete oppure crea un censimento nuovo." }
    if (-not $Cliente -and -not $NonChiedere) { $Cliente = (Read-Host 'Nuovo censimento: nome del cliente').Trim() }
    (New-Censimento -Cliente $Cliente -Confirm:$false).FullName
}

function Save-HDCensimento {
    # Scrittura sicura: copia temporanea, modifica, rilettura di controllo, backup, sostituzione.
    # -Extra: modifiche per gli altri fogli (@{ Tabella; Modifiche }) oppure fogli da riscrivere
    # per intero (@{ Tabella; Riscrivi = righe }). I fogli che mancano vengono creati.
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][hashtable]$Modifiche, [object[]]$Extra = @())
    $lista = @(ConvertTo-HDModificheXlsx -Tabella $Tabella -Modifiche $Modifiche)
    $altri = @(foreach ($e in @($Extra | Where-Object { $_ -and $_.Tabella })) {
            if ($null -ne $e.Riscrivi) { [pscustomobject]@{ Tabella = $e.Tabella; Lista = @(); Riscrivi = @($e.Riscrivi) }; continue }
            $l = @(ConvertTo-HDModificheXlsx -Tabella $e.Tabella -Modifiche $e.Modifiche)
            if ($l.Count -gt 0 -and $e.Tabella.ScriviIntestazione) {
                $celle = @{}
                for ($i = 0; $i -lt $e.Tabella.Definizione.Count; $i++) { $celle[$i + 1] = $e.Tabella.Definizione[$i].Titolo }
                $l = @(@{ Riga = 1; Celle = $celle }) + $l
            }
            if ($l.Count -gt 0) { [pscustomobject]@{ Tabella = $e.Tabella; Lista = $l; Riscrivi = $null } }
        })
    if ($lista.Count -eq 0 -and $altri.Count -eq 0) { return [pscustomobject]@{ Scritto = $false; Backup = $null; Saltate = @() } }
    $cfg = Get-HDConfig
    $p = $Tabella.Percorso
    $cartella = Split-Path -Parent $p
    $nome = [System.IO.Path]::GetFileNameWithoutExtension($p)
    $temporaneo = Join-Path $cartella ('.~{0}.tmp.xlsx' -f $nome)
    $backup = $null
    $numeriche = @($script:HDColonneCensimento | Where-Object { $_.Numerica -and $Tabella.Colonne.ContainsKey($_.Chiave) } | ForEach-Object { [int]$Tabella.Colonne[$_.Chiave] })
    try {
        Copy-Item -LiteralPath $p -Destination $temporaneo -Force -ErrorAction Stop
        $saltate = New-Object System.Collections.Generic.List[string]
        if ($lista.Count -gt 0) {
            $celleSaltate = Update-HDXlsx -Percorso $temporaneo -ParteFoglio $Tabella.ParteFoglio -IndiceFoglio $Tabella.IndiceFoglio -Modifiche $lista -ColonneNumeriche $numeriche -RigaIntestazione $Tabella.RigaIntestazione
            foreach ($x in $celleSaltate) { if ($x) { $saltate.Add($x) } }
        }
        foreach ($a in $altri) {
            $parte = $a.Tabella.ParteFoglio
            $indice = $a.Tabella.IndiceFoglio
            if (-not $a.Tabella.Esiste) {
                $creato = Add-HDFoglioXlsx -Percorso $temporaneo -NomeFoglio $a.Tabella.Nome -Colonne $a.Tabella.Definizione -Nascosto:([bool]$a.Tabella.Nascosto)
                $parte = $creato.Parte
                $indice = $creato.Indice
            }
            if ($null -ne $a.Riscrivi) {
                $numericheFoglio = @(for ($i = 0; $i -lt $a.Tabella.Definizione.Count; $i++) { if ($a.Tabella.Definizione[$i].Numerica) { $i + 1 } })
                Set-HDRigheFoglio -Percorso $temporaneo -ParteFoglio $parte -IndiceFoglio $indice -Intestazione @($a.Tabella.Definizione | ForEach-Object { $_.Titolo }) -Righe $a.Riscrivi -ColonneNumeriche $numericheFoglio
                continue
            }
            $celleSaltate = Update-HDXlsx -Percorso $temporaneo -ParteFoglio $parte -IndiceFoglio $indice -Modifiche $a.Lista -RigaIntestazione $a.Tabella.RigaIntestazione
            foreach ($x in $celleSaltate) { if ($x) { $saltate.Add($x) } }
        }
        $verifica = ConvertTo-HDTabellaCensimento -Fogli (Get-HDFogliPrincipali @(Read-HDXlsx -Percorso $temporaneo -EscludiFogli @($script:HDFoglioSoftware))) -Percorso $temporaneo
        if ($verifica.Righe.Count -lt $Tabella.RigheIniziali) { throw 'controllo fallito dopo la scrittura (righe mancanti): Excel non modificato.' }
        $cartellaBackup = Join-Path $cartella 'backup'
        [void][System.IO.Directory]::CreateDirectory($cartellaBackup)
        $backup = Join-Path $cartellaBackup ('{0}_{1}.xlsx' -f $nome, (Get-Date).ToString('yyyyMMdd-HHmmss', $script:HDInvariante))
        Copy-Item -LiteralPath $p -Destination $backup -Force -ErrorAction Stop
        Remove-Item -LiteralPath $p -Force -ErrorAction Stop
        Move-Item -LiteralPath $temporaneo -Destination $p -ErrorAction Stop
        Get-ChildItem -LiteralPath $cartellaBackup -Filter ($nome + '_*.xlsx') -File |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip ([int]$cfg.Censimento.BackupDaTenere) |
            Remove-Item -Force -ErrorAction SilentlyContinue
        [pscustomobject]@{ Scritto = $true; Backup = $backup; Saltate = @($saltate.ToArray()) }
    }
    catch {
        $errore = $_.Exception
        if (Test-Path -LiteralPath $temporaneo) { Remove-Item -LiteralPath $temporaneo -Force -ErrorAction SilentlyContinue }
        if (-not (Test-Path -LiteralPath $p) -and $backup -and (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $backup -Destination $p -ErrorAction SilentlyContinue }
        if ($errore -is [System.IO.IOException]) { throw "Non riesco a scrivere nell'Excel, forse è aperto in Excel: chiudilo e riprova. ($($errore.Message))" }
        throw
    }
}

#endregion

function Get-HDFileScansioneRecenti {
    # Ultimi CSV/XML nella cartella del censimento, sul Desktop, in Documenti e in Download
    $cartelle = @(Get-HDCartella -Tipo Censimento)
    foreach ($c in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('MyDocuments'))) { if ($c) { $cartelle += $c } }
    if ($env:USERPROFILE) { $cartelle += (Join-Path $env:USERPROFILE 'Downloads') }
    @($cartelle | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ } |
            ForEach-Object { Get-ChildItem -LiteralPath $_ -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.csv', '.xml' } } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 9)
}

function Select-HDFileScansione {
    $file = @(Get-HDFileScansioneRecenti)
    if ($file.Count -gt 0) {
        Write-Host 'File CSV/XML recenti (Censimento, Desktop, Documenti, Download):'
        for ($i = 0; $i -lt $file.Count; $i++) {
            Write-Host ('  {0}  {1}   ({2}, {3})' -f ($i + 1), $file[$i].Name, (Split-Path -Leaf $file[$i].DirectoryName), (Format-HDData $file[$i].LastWriteTime))
        }
    }
    $scelta = (Read-Host 'Numero del file, oppure trascina qui il file della scansione e premi INVIO').Trim().Trim('"')
    if (-not $scelta) { return $null }
    if ($scelta -match '^\d$' -and [int]$scelta -ge 1 -and [int]$scelta -le $file.Count) { return $file[[int]$scelta - 1].FullName }
    $scelta
}

function Show-HDEsitoImport {
    param([Parameter(Mandatory)]$Esito)
    Write-Host ''
    Write-Host (' Scansione importata ({0})' -f $Esito.Formato) -ForegroundColor Green
    Write-Host ('  Dispositivi attivi: {0}   aggiunti: {1}   aggiornati: {2}   invariati: {3}   scartati (spenti o vuoti): {4}' -f $Esito.Dispositivi, $Esito.Aggiunti, $Esito.Aggiornati, $Esito.Invariati, $Esito.Scartati)
    Write-Host ('  Excel: {0}' -f $Esito.Censimento)
    if (@($Esito.ColonneMancanti).Count -gt 0) { Write-Host ('  Colonne non trovate nell''Excel: {0}' -f (@($Esito.ColonneMancanti) -join ', ')) -ForegroundColor Yellow }
    Write-Host '  Ora porta la chiavetta sui PC e lancia Censisci-PC.cmd: completerà le righe.' -ForegroundColor Cyan
}

#region Rilevazioni (copia di sicurezza di ogni censimento)

function Get-HDCartellaRilevazioni {
    param([Parameter(Mandatory)][string]$FileCensimento, [switch]$InAttesa)
    $c = Join-Path (Join-Path (Split-Path -Parent $FileCensimento) 'rilevazioni') ([System.IO.Path]::GetFileNameWithoutExtension($FileCensimento))
    if ($InAttesa) { $c = Join-Path $c 'da-inserire' }
    $c
}

function Save-HDRilevazione {
    param([Parameter(Mandatory)]$Record, [hashtable]$Manuali, [Parameter(Mandatory)][string]$FileCensimento)
    $cartella = Get-HDCartellaRilevazioni -FileCensimento $FileCensimento -InAttesa
    [void][System.IO.Directory]::CreateDirectory($cartella)
    $nome = Get-HDPercorsoUnico -Cartella $cartella -NomeBase ('{0}_{1}' -f (ConvertTo-HDNomeFile $Record.Hostname), (Get-Date).ToString('yyyyMMdd-HHmmss', $script:HDInvariante)) -Estensione '.json'
    $dati = [pscustomobject]@{
        Censimento = (Split-Path -Leaf $FileCensimento)
        Rilevato   = (Get-Date).ToString('o', $script:HDInvariante)
        Dati       = ($Record | Select-Object -Property * -ExcludeProperty Rilevato)
        Manuali    = $Manuali
    }
    Write-HDFileTesto -Percorso $nome -Testo ($dati | ConvertTo-Json -Depth 5)
    $nome
}

function Read-HDRilevazione {
    param([Parameter(Mandatory)][string]$Percorso)
    $j = Read-HDFileTesto $Percorso | ConvertFrom-Json
    $manuali = @{}
    if ($j.Manuali) { foreach ($p in $j.Manuali.PSObject.Properties) { $manuali[$p.Name] = [string]$p.Value } }
    [pscustomobject]@{ File = $Percorso; Dati = $j.Dati; Manuali = $manuali; Rilevato = $j.Rilevato }
}

function ConvertTo-HDTabellaHash {
    # Hashtable oppure oggetto letto dal JSON -> hashtable
    param($Oggetto)
    if ($Oggetto -is [hashtable]) { return $Oggetto }
    $h = @{}
    if ($Oggetto) { foreach ($p in $Oggetto.PSObject.Properties) { $h[$p.Name] = $p.Value } }
    $h
}

function Add-HDExtraATabelle {
    # Rilevazioni -> modifiche per i fogli Dettagli e Software
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][object[]]$Rilevazioni)
    $modDettagli = @{}
    $modDati = @{}
    $riepilogo = @{}
    foreach ($nome in 'Dettagli', 'Software', 'DatiSoftware') {
        $t = $Tabella.$nome
        if ($t -and $t.Conflitto) { Write-Warning ("Nell'Excel c'è già un foglio '{0}' con un altro contenuto: quei dati non vengono scritti (rinomina il tuo foglio per averli)." -f $t.Nome) }
    }
    $conDettagli = $Tabella.Dettagli -and -not $Tabella.Dettagli.Conflitto
    $conSoftware = $Tabella.DatiSoftware -and -not $Tabella.DatiSoftware.Conflitto -and $Tabella.Software -and -not $Tabella.Software.Conflitto
    if ($conDettagli -and $Tabella.Dettagli.Esiste) { Add-HDColonneMancanti -Tabella $Tabella.Dettagli -Colonne $Tabella.Dettagli.DaAggiungere -Modifiche $modDettagli }
    foreach ($ril in $Rilevazioni) {
        $rec = $ril.Dati
        $data = ''
        if ($ril.Rilevato -is [datetime]) { $data = Format-HDData $ril.Rilevato }
        elseif ($ril.Rilevato) {
            $quando = [datetime]::MinValue
            if ([datetime]::TryParse([string]$ril.Rilevato, $script:HDInvariante, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$quando)) { $data = Format-HDData $quando.ToLocalTime() }
        }
        if ($rec.Dettagli -and $conDettagli) {
            $valori = @{}
            $grezzi = ConvertTo-HDTabellaHash $rec.Dettagli
            foreach ($k in $grezzi.Keys) { $valori[$k] = [string]$grezzi[$k] }
            $valori['Hostname'] = [string]$rec.Hostname
            $valori['SerialNumber'] = [string]$rec.SerialNumber
            $valori['Rilevato'] = $data
            [void](Add-HDDettagliATabella -Tabella $Tabella.Dettagli -Valori $valori -Modifiche $modDettagli -SeVuoti @{ Smart = 'Avvia il toolkit come amministratore per leggerlo' })
        }
        if ($null -ne $rec.Software -and $conSoftware) {
            $programmi = @($rec.Software | Where-Object { $_ })
            $valori = @{ Hostname = [string]$rec.Hostname; SerialNumber = [string]$rec.SerialNumber; Rilevato = $data; Programmi = (ConvertTo-HDTestoProgrammi $programmi) }
            [void](Add-HDDettagliATabella -Tabella $Tabella.DatiSoftware -Valori $valori -Modifiche $modDati)
            $riepilogo[[string]$rec.Hostname] = [pscustomobject]@{ Totale = $programmi.Count }
        }
    }
    $extra = @(@{ Tabella = $Tabella.Dettagli; Modifiche = $modDettagli })
    if ($riepilogo.Count -gt 0) {
        # il foglio Software si ricostruisce da DatiSoftware (già aggiornato in memoria)
        $extra += @{ Tabella = $Tabella.Software; Riscrivi = @(ConvertTo-HDRigheSoftware $Tabella.DatiSoftware.Righe) }
        $extra += @{ Tabella = $Tabella.DatiSoftware; Modifiche = $modDati }
    }
    [pscustomobject]@{ Extra = $extra; Riepilogo = $riepilogo }
}

function Invoke-HDSincronizzazione {
    # Inserisce nell'Excel tutte le rilevazioni in attesa (dalla più vecchia)
    param([Parameter(Mandatory)][string]$FileCensimento)
    $cartella = Get-HDCartellaRilevazioni -FileCensimento $FileCensimento -InAttesa
    $file = @(Get-ChildItem -LiteralPath $cartella -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object { ($_.BaseName -split '_')[-1] }, Name)
    $vuoto = [pscustomobject]@{ Esiti = @(); Scritto = $false; Backup = $null; Saltate = @(); Errore = $null; Inserite = 0; Software = @{} }
    if ($file.Count -eq 0) { return $vuoto }
    $rilevazioni = @(foreach ($f in $file) {
            try { Read-HDRilevazione $f.FullName }
            catch { Write-Warning "Rilevazione illeggibile, la lascio da parte: $($f.FullName)" }
        })
    if ($rilevazioni.Count -eq 0) { return $vuoto }
    $tabella = Get-HDTabellaCensimento $FileCensimento
    $applicate = Add-HDRilevazioniATabella -Tabella $tabella -Rilevazioni $rilevazioni
    # colonne arrivate con le versioni nuove (es. BitLocker): si aggiungono in fondo se servono
    $daAggiungere = @($tabella.DaAggiungere | Where-Object { $k = $_.Chiave; @($rilevazioni | Where-Object { [string]$_.Dati.$k }).Count -gt 0 })
    if ($daAggiungere.Count -gt 0) { Add-HDColonneMancanti -Tabella $tabella -Colonne $daAggiungere -Modifiche $applicate.Modifiche }
    $extra = Add-HDExtraATabelle -Tabella $tabella -Rilevazioni $rilevazioni
    try { $s = Save-HDCensimento -Tabella $tabella -Modifiche $applicate.Modifiche -Extra $extra.Extra }
    catch { return [pscustomobject]@{ Esiti = $applicate.Esiti; Scritto = $false; Backup = $null; Saltate = @(); Errore = $_.Exception.Message; Inserite = 0; Software = $extra.Riepilogo } }
    $archivio = Split-Path -Parent $cartella
    foreach ($r in $rilevazioni) { Move-Item -LiteralPath $r.File -Destination (Join-Path $archivio (Split-Path -Leaf $r.File)) -Force -ErrorAction SilentlyContinue }
    [pscustomobject]@{ Esiti = $applicate.Esiti; Scritto = $s.Scritto; Backup = $s.Backup; Saltate = $s.Saltate; Errore = $null; Inserite = $rilevazioni.Count; Software = $extra.Riepilogo }
}

#endregion

#region Comandi

function New-Censimento {
    <#
    .SYNOPSIS
    Crea un nuovo file Excel di censimento con le colonne standard.

    .DESCRIPTION
    Il file nasce nella cartella Censimento del toolkit (quindi sulla chiavetta) con
    intestazione bloccata, filtri, colonne da compilare a mano evidenziate in giallo,
    elenchi a discesa per Tipo IP e Cynet, Cynet "No" in rosso. Ha anche i fogli Dettagli
    (Windows 11, dischi, monitor, stampanti, Office, licenza...) e Software (programmi
    installati). Se nella cartella c'è un tuo Modello.xlsx viene usato quello: i fogli
    Dettagli e Software si aggiungono da soli al primo censimento.

    .EXAMPLE
    New-Censimento 'Studio Rossi'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Position = 0)][string]$Cliente, [string]$Percorso)
    if (-not $Percorso) {
        $nome = 'Censimento'
        if ($Cliente) { $nome = 'Censimento_' + (ConvertTo-HDNomeFile $Cliente) }
        $Percorso = Get-HDPercorsoUnico -Cartella (Get-HDCartella -Tipo Censimento) -NomeBase $nome -Estensione '.xlsx'
    }
    $p = Resolve-HDPercorso $Percorso
    if (Test-Path -LiteralPath $p) { throw "Il file esiste già: $p" }
    if (-not $PSCmdlet.ShouldProcess($p, 'Crea il file di censimento')) { return }
    $modello = Get-HDModelloCensimento
    if ($modello) {
        [void][System.IO.Directory]::CreateDirectory((Split-Path -Parent $p))
        Copy-Item -LiteralPath $modello -Destination $p
        [void](Get-HDTabellaCensimento $p)
    }
    else {
        $indice = @{}
        for ($i = 0; $i -lt $script:HDColonneCensimento.Count; $i++) { $indice[$script:HDColonneCensimento[$i].Chiave] = $i + 1 }
        $dettagli = @{}
        for ($i = 0; $i -lt $script:HDColonneDettagli.Count; $i++) { $dettagli[$script:HDColonneDettagli[$i].Chiave] = $i + 1 }
        $altri = @(
            @{ Nome = $script:HDFoglioDettagli; Colonne = $script:HDColonneDettagli; Evidenzia = @{
                    $dettagli['Windows11']          = 'LEFT({0},3)="No:"'
                    $dettagli['DischiFisici']       = 'OR(ISNUMBER(SEARCH("GUASTO",{0})),ISNUMBER(SEARCH("ATTENZIONE",{0})))'
                    $dettagli['Smart']              = 'OR(ISNUMBER(SEARCH("GUASTO",{0})),ISNUMBER(SEARCH("ATTENZIONE",{0})))'
                    $dettagli['AttivazioneWindows'] = 'LEFT({0},12)="Non attivato"'
                }
            }
            @{ Nome = $script:HDFoglioSoftware; Colonne = $script:HDColonneSoftware }
            @{ Nome = $script:HDFoglioDatiSoftware; Colonne = $script:HDColonneDatiSoftware; Nascosto = $true }
        )
        New-HDXlsx -Percorso $p -NomeFoglio 'Censimento' -Colonne $script:HDColonneCensimento `
            -Elenchi @{ $indice['TipoIP'] = 'DHCP,Statico'; $indice['Cynet'] = 'Sì,No' } `
            -Evidenzia @{ $indice['Cynet'] = '{0}="No"' } -AltriFogli $altri
    }
    Get-Item -LiteralPath $p
}

function Import-HDDispositiviNelCensimento {
    # Dispositivi (da file di Advanced IP Scanner o dalla scansione integrata) -> righe dell'Excel
    param(
        [Parameter(Mandatory)][object[]]$Dispositivi,
        [int]$Scartati = 0,
        [string]$Formato,
        [string]$Censimento,
        [string]$Cliente,
        [switch]$NonChiedere,
        [bool]$Scrivi = $true
    )
    $interattivo = (-not $NonChiedere) -and (Test-HDInterattivo)
    $file = Resolve-HDFileCensimento -Percorso $Censimento -Cliente $Cliente -Crea -NonChiedere:(-not $interattivo)
    $tabella = Get-HDTabellaCensimento $file
    $applicate = Add-HDScansioneATabella -Tabella $tabella -Dispositivi $Dispositivi
    $salvataggio = $null
    if ($Scrivi) { $salvataggio = Save-HDCensimento -Tabella $tabella -Modifiche $applicate.Modifiche }
    [pscustomobject]@{
        PSTypeName      = 'HelpDesk.ImportScansione'
        Censimento      = $file
        Formato         = $Formato
        Dispositivi     = @($Dispositivi).Count
        Aggiunti        = @($applicate.Esiti | Where-Object Azione -eq 'Aggiunto').Count
        Aggiornati      = @($applicate.Esiti | Where-Object Azione -eq 'Aggiornato').Count
        Invariati       = @($applicate.Esiti | Where-Object Azione -eq 'Invariato').Count
        Scartati        = $Scartati
        Backup          = $(if ($salvataggio) { $salvataggio.Backup })
        ColonneMancanti = $tabella.Mancanti
        Esiti           = $applicate.Esiti
    }
}

function Import-ScansioneRete {
    <#
    .SYNOPSIS
    Inserisce nel censimento Excel i dispositivi trovati da Advanced IP Scanner.

    .DESCRIPTION
    In Advanced IP Scanner esporta i risultati in CSV (o XML), poi passa il file a questo
    comando (o trascinalo su Importa-Scansione.cmd). Per ogni dispositivo attivo:
    se è già nel censimento (stesso MAC o hostname) ne aggiorna l'indirizzo IP e completa i
    campi vuoti, altrimenti aggiunge una riga con Hostname, Indirizzo IP, MAC Address e
    Manufacturer (il produttore della scheda di rete: verrà corretto dal censimento sul PC).
    Le colonne che compili tu (Utilizzatore, Note, Locazione) non vengono mai toccate.
    Se non esiste ancora un censimento lo crea. In alternativa c'è Invoke-ScansioneRete,
    che fa la scansione da solo.

    .EXAMPLE
    Import-ScansioneRete .\Censimento\scansione.csv -Cliente 'Studio Rossi'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position = 0)][Alias('FullName')][string]$Percorso,
        [string]$Censimento,
        [string]$Cliente,
        [switch]$NonChiedere
    )
    $p = Resolve-HDPercorso $Percorso
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { throw "File della scansione non trovato: $Percorso" }
    if ($p -match '\.(html?|xlsx?)$') { throw 'Da Advanced IP Scanner esporta la scansione in formato CSV (oppure XML) e importa quel file.' }
    $scansione = ConvertFrom-HDScansione -Testo (Read-HDFileTesto $p)
    if ($scansione.Dispositivi.Count -eq 0) { throw "Nella scansione non ci sono dispositivi attivi ($($scansione.Ignorati) righe scartate perché spenti o vuoti)." }
    $scrivi = $PSCmdlet.ShouldProcess($p, "Importa $($scansione.Dispositivi.Count) dispositivi nel censimento")
    Import-HDDispositiviNelCensimento -Dispositivi $scansione.Dispositivi -Scartati $scansione.Ignorati -Formato $scansione.Formato -Censimento $Censimento -Cliente $Cliente -NonChiedere:$NonChiedere -Scrivi $scrivi
}

function Show-HDSchedaCensimento {
    param([Parameter(Mandatory)]$Record, $Trovata, [string]$File, [int]$SogliaDisco = 10)
    Write-Host ''
    Write-Host (' CENSIMENTO PC  {0} ' -f $Record.Hostname) -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host (' Excel: {0}' -f $File) -ForegroundColor DarkGray
    if ($Trovata) { Write-Host (' Già nel censimento alla riga {0} ({1}): verrà aggiornato' -f $Trovata.Riga.Numero, $Trovata.Motivo) -ForegroundColor Cyan }
    else { Write-Host ' Non è nel censimento: verrà aggiunto in fondo' -ForegroundColor Cyan }
    $ip = $Record.IndirizzoIP
    if ($ip) { $ip = '{0} ({1})' -f $ip, $Record.TipoIP } else { $ip = '(non collegato alla rete)' }
    $utente = $Record.NomeUtente
    if ($Record.FonteNomeUtente) { $utente = '{0}   ({1})' -f $utente, $Record.FonteNomeUtente }
    $seriale = $Record.SerialNumber
    if (-not $seriale) {
        $seriale = '(non presente nel BIOS)'
        if ($Record.SerialeBios) { $seriale = "(nel BIOS c'è solo '$($Record.SerialeBios)')" }
    }
    $righe = [ordered]@{
        'Indirizzo IP'      = $ip
        'MAC Address'       = $Record.MAC
        'Device Type'       = $Record.DeviceType
        'Manufacturer'      = $Record.Manufacturer
        'Model'             = $Record.Model
        'Serial Number'     = $seriale
        'Nome Utente'       = $utente
        'Dominio'           = $Record.Dominio
        'Operating System'  = ('{0}  {1}' -f $Record.OS, $Record.OSVersion).Trim()
        'Memory'            = $Record.Memory
        'Processor Type'    = $Record.Processor
        'Free Disk Space'   = $Record.FreeDiskPct
        'Logical Disk Size' = $Record.DiskSize
        'BitLocker'         = $Record.BitLocker
    }
    foreach ($k in $righe.Keys) { Write-Host ('  {0,-19}{1}' -f $k, $righe[$k]) }
    if ($Record.Cynet -eq 'Sì') {
        Write-Host ('  {0,-19}{1}' -f 'Cynet', ('Sì  (' + $Record.CynetDettaglio + ')')) -ForegroundColor Green
        if (-not $Record.CynetInEsecuzione) { Write-Host '                     Attenzione: Cynet risulta installato ma l''agente non è in esecuzione.' -ForegroundColor Yellow }
    }
    else { Write-Host ('  {0,-19}{1}' -f 'Cynet', 'No') -ForegroundColor Red }
    if ($null -ne $Record.DiscoMinimoPct -and "$($Record.DiscoMinimoPct)" -ne '' -and [int]$Record.DiscoMinimoPct -lt $SogliaDisco) { Write-Host '  Poco spazio libero su un disco.' -ForegroundColor Yellow }
    if ($Record.Dettagli) {
        $d = ConvertTo-HDTabellaHash $Record.Dettagli
        Write-Host ' Dettagli (foglio Dettagli):' -ForegroundColor White
        foreach ($c in $script:HDColonneDettagli) {
            if ($c.Chiave -in 'Hostname', 'SerialNumber', 'Rilevato' -or -not [string]$d[$c.Chiave]) { continue }
            $colore = 'Gray'
            if ($c.Chiave -eq 'Windows11') {
                if ($d[$c.Chiave] -match '^(Sì|Già)') { $colore = 'Green' } elseif ($d[$c.Chiave] -match '^No') { $colore = 'Red' } else { $colore = 'Yellow' }
            }
            elseif ([string]$d[$c.Chiave] -match 'GUASTO|ATTENZIONE|^Non attivato') { $colore = 'Red' }
            Write-Host ('  {0,-19}{1}' -f $c.Titolo, $d[$c.Chiave]) -ForegroundColor $colore
        }
    }
    Write-Host ''
}

function Read-HDCampoManuale {
    param([string]$Etichetta, [string]$Attuale)
    $suggerimento = ''
    if ($Attuale) { $suggerimento = " [$Attuale]" }
    (Read-Host ('  {0}{1}' -f $Etichetta, $suggerimento)).Trim()
}

function Add-PCCensimento {
    <#
    .SYNOPSIS
    Censisce questo PC e lo scrive nell'Excel del censimento (da lanciare dalla chiavetta).

    .DESCRIPTION
    Raccoglie i dati del PC, li mostra, chiede (facoltativi) Utilizzatore, Locazione e Note,
    poi aggiorna la riga del PC nell'Excel: la ritrova per MAC, numero di serie o hostname.
    Se il PC non c'è (per esempio era spento o fuori rete durante la scansione) lo aggiunge.
    Prima di scrivere fa un backup dell'Excel; ogni rilevazione è salvata anche in
    rilevazioni\ così, se l'Excel è aperto o la scrittura non riesce, nulla va perso e i dati
    vengono inseriti al censimento successivo (o con Sync-Censimento).

    .EXAMPLE
    Add-PCCensimento

    .PARAMETER Record
    Dati del PC già letti con Get-InventarioPC (la finestra li passa per non rileggerli).

    .EXAMPLE
    Add-PCCensimento -Utilizzatore 'Mario Rossi' -Locazione 'Reception' -NonChiedere
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$Censimento,
        [object[]]$Record,
        [string[]]$ComputerName,
        [pscredential]$Credential,
        [string]$Utilizzatore,
        [string]$Locazione,
        [string]$Note,
        [string]$NomeUtente,
        [switch]$NonChiedere
    )
    $cfg = Get-HDConfig
    $interattivo = (-not $NonChiedere) -and (-not $ComputerName) -and (Test-HDInterattivo)
    $file = Resolve-HDFileCensimento -Percorso $Censimento -Crea -NonChiedere:(-not $interattivo)
    $tabella = Get-HDTabellaCensimento $file
    if ($tabella.Mancanti.Count -gt 0) { Write-Warning ("Nell'Excel mancano le colonne: {0}. Quei dati non verranno scritti." -f ($tabella.Mancanti -join ', ')) }

    if ($Record) { $raccolti = @($Record | Where-Object { $_ }) }
    else { $raccolti = @(Get-InventarioPC -ComputerName $ComputerName -Credential $Credential) }
    if ($raccolti.Count -eq 0) { throw 'Nessun dato raccolto: censimento non eseguito.' }
    foreach ($r in $raccolti) {
        $trovata = Find-HDRigaCensimento -Righe $tabella.Righe -Mac (@($r.TuttiMAC) + @($r.MAC)) -Seriale $r.SerialNumber -Hostname $r.Hostname
        Show-HDSchedaCensimento -Record $r -Trovata $trovata -File $file -SogliaDisco ([int]$cfg.Soglie.DiscoAttenzionePct)
        $attuali = @{}
        if ($trovata) { $attuali = $trovata.Riga.Valori }
        $manuali = @{}
        if ($Utilizzatore) { $manuali['Utilizzatore'] = $Utilizzatore }
        if ($Locazione) { $manuali['Locazione'] = $Locazione }
        if ($Note) { $manuali['Note'] = $Note }
        if ($NomeUtente) { $manuali['NomeUtente'] = $NomeUtente }
        if ($interattivo) {
            Write-Host ' Completa (INVIO = lascia com''è):' -ForegroundColor White
            foreach ($campo in @(@{ K = 'Utilizzatore'; E = 'Utilizzatore' }, @{ K = 'Locazione'; E = 'Locazione' }, @{ K = 'Note'; E = 'Note' })) {
                if ($manuali.ContainsKey($campo.K)) { continue }
                $v = Read-HDCampoManuale -Etichetta $campo.E -Attuale $attuali[$campo.K]
                if ($v) { $manuali[$campo.K] = $v }
            }
            if ($r.FonteNomeUtente -ne 'utente collegato' -and -not $manuali.ContainsKey('NomeUtente')) {
                $v = Read-HDCampoManuale -Etichetta 'Nome Utente (correggi se serve)' -Attuale $r.NomeUtente
                if ($v) { $manuali['NomeUtente'] = $v }
            }
        }
        if ($PSCmdlet.ShouldProcess($file, "Censisci $($r.Hostname)")) {
            [void](Save-HDRilevazione -Record $r -Manuali $manuali -FileCensimento $file)
        }
    }
    if ($WhatIfPreference) { return }

    $esito = Invoke-HDSincronizzazione -FileCensimento $file
    foreach ($e in $esito.Esiti) {
        $campi = 'nessun dato cambiato'
        if ($e.Campi.Count -gt 0) { $campi = 'campi scritti: ' + ($e.Campi -join ', ') }
        if ($esito.Errore) { continue }
        if ($e.Azione -eq 'Aggiunto') { Write-Host (' {0}: aggiunto alla riga {1} ({2})' -f $e.Hostname, $e.Riga, $campi) -ForegroundColor Green }
        else { Write-Host (' {0}: riga {1} aggiornata, {2} ({3})' -f $e.Hostname, $e.Riga, $e.Motivo, $campi) -ForegroundColor Green }
    }
    if ($esito.Errore) {
        Write-Host (" Excel NON aggiornato: {0}" -f $esito.Errore) -ForegroundColor Red
        Write-Host (" I dati sono al sicuro in {0}: verranno inseriti al prossimo censimento (o con Sync-Censimento)." -f (Get-HDCartellaRilevazioni -FileCensimento $file -InAttesa)) -ForegroundColor Yellow
    }
    elseif (@($esito.Saltate).Count -gt 0) { Write-Warning ('Celle con formule lasciate com''erano: ' + (@($esito.Saltate) -join ', ')) }
    if (-not $esito.Errore) {
        foreach ($h in @($esito.Software.Keys)) {
            $sw = $esito.Software[$h]
            Write-Host (' {0}: {1} programmi, riepilogati nel foglio Software' -f $h, $sw.Totale) -ForegroundColor Green
        }
    }
    if ($esito.Inserite -gt $raccolti.Count) { Write-Host (' Inserite anche {0} rilevazioni rimaste in sospeso.' -f ($esito.Inserite - $raccolti.Count)) -ForegroundColor Cyan }
    [pscustomobject]@{
        PSTypeName = 'HelpDesk.EsitoCensimento'
        Censimento = $file
        Scritto    = (-not $esito.Errore)
        Esiti      = $esito.Esiti
        Backup     = $esito.Backup
        Errore     = $esito.Errore
    }
}

function Sync-Censimento {
    <#
    .SYNOPSIS
    Inserisce nell'Excel le rilevazioni rimaste in sospeso (es. il file era aperto in Excel).

    .EXAMPLE
    Sync-Censimento
    #>
    [CmdletBinding()]
    param([string]$Censimento, [switch]$NonChiedere)
    $file = Resolve-HDFileCensimento -Percorso $Censimento -NonChiedere:($NonChiedere -or -not (Test-HDInterattivo))
    $esito = Invoke-HDSincronizzazione -FileCensimento $file
    if ($esito.Errore) { throw $esito.Errore }
    [pscustomobject]@{ Censimento = $file; Inserite = $esito.Inserite; Esiti = $esito.Esiti; Backup = $esito.Backup }
}

#endregion

# Censimento PC: colonne dell'Excel, lettura della scansione di rete (Advanced IP Scanner),
# abbinamento delle righe e conversione dei dati raccolti sul PC.
# Funzioni pure: nessuna dipendenza da Windows, coperte dai test.

function New-HDColonnaCensimento {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param([string]$Chiave, [string]$Titolo, [double]$Larghezza, [string[]]$Sinonimi, [switch]$Manuale, [switch]$Numerica, [switch]$Aggiungibile)
    [pscustomobject]@{ Chiave = $Chiave; Titolo = $Titolo; Larghezza = $Larghezza; Sinonimi = $Sinonimi; Manuale = [bool]$Manuale; Numerica = [bool]$Numerica; Testo = (-not $Numerica); Aggiungibile = [bool]$Aggiungibile }
}

# Colonne nell'ordine dell'Excel. I sinonimi servono a riconoscere anche file con intestazioni diverse.
$script:HDColonneCensimento = @(
    New-HDColonnaCensimento 'Hostname' 'Hostname' 18 @('hostname', 'host name', 'nome host', 'nome pc', 'computer', 'nome computer')
    New-HDColonnaCensimento 'IndirizzoIP' 'Indirizzo IP' 15 @('indirizzo ip', 'ip', 'ip address', 'indirizzo')
    New-HDColonnaCensimento 'TipoIP' 'Tipo IP' 9 @('tipo ip', 'tipo indirizzo', 'dhcp')
    New-HDColonnaCensimento 'MAC' 'MAC Address' 19 @('mac address', 'mac', 'indirizzo mac', 'mac-address')
    New-HDColonnaCensimento 'DeviceType' 'Device Type' 13 @('device type', 'tipo dispositivo', 'tipo device', 'tipologia', 'tipo')
    New-HDColonnaCensimento 'Manufacturer' 'Manufacturer' 14 @('manufacturer', 'produttore', 'marca', 'vendor', 'costruttore')
    New-HDColonnaCensimento 'Model' 'Model' 26 @('model', 'modello')
    New-HDColonnaCensimento 'NomeUtente' 'Nome Utente' 16 @('nome utente', 'utente', 'username', 'user name', 'account')
    New-HDColonnaCensimento 'Utilizzatore' 'Utilizzatore' 22 @('utilizzatore', 'assegnatario', 'assegnato a') -Manuale
    New-HDColonnaCensimento 'Note' 'Note' 30 @('note', 'nota', 'notes') -Manuale
    New-HDColonnaCensimento 'Dominio' 'Dominio' 18 @('dominio', 'domain', 'workgroup', 'dominio workgroup')
    New-HDColonnaCensimento 'SerialNumber' 'Serial Number' 16 @('serial number', 'numero di serie', 'seriale', 's n', 'sn', 'serial')
    New-HDColonnaCensimento 'OS' 'Operating System' 22 @('operating system', 'sistema operativo', 'os')
    New-HDColonnaCensimento 'OSVersion' 'Operating System Version' 19 @('operating system version', 'versione sistema operativo', 'versione os', 'os version', 'versione')
    New-HDColonnaCensimento 'Memory' 'Memory' 26 @('memory', 'memoria', 'ram', 'memoria ram')
    New-HDColonnaCensimento 'Processor' 'Processor Type' 36 @('processor type', 'processore', 'cpu', 'processor', 'tipo processore')
    New-HDColonnaCensimento 'FreeDiskPct' 'Free Disk Space %' 30 @('free disk space', 'spazio libero', 'spazio disco libero', 'free disk', 'disco libero')
    New-HDColonnaCensimento 'DiskSize' 'Logical Disk Size' 24 @('logical disk size', 'dimensione disco', 'disk size', 'capacita disco')
    New-HDColonnaCensimento 'Cynet' 'Cynet' 8 @('cynet')
    # Colonne arrivate dopo: negli Excel che non le hanno vengono aggiunte in fondo
    New-HDColonnaCensimento 'BitLocker' 'BitLocker' 20 @('bitlocker', 'crittografia', 'cifratura disco') -Aggiungibile
    New-HDColonnaCensimento 'Locazione' 'Locazione' 22 @('locazione', 'ubicazione', 'posizione', 'ufficio', 'stanza', 'location') -Manuale
)

# Colonne aggiornate dallo script sul PC (le manuali solo se le scrivi tu quando richiesto)
$script:HDColonneAutomatiche = @('Hostname', 'DeviceType', 'Manufacturer', 'Model', 'NomeUtente', 'Dominio', 'SerialNumber', 'OS', 'OSVersion', 'Memory', 'Processor', 'FreeDiskPct', 'DiskSize', 'Cynet', 'BitLocker')

$script:HDIpv4 = '^\d{1,3}(\.\d{1,3}){3}$'
$script:HDSchedeVirtuali = '(?i)hyper-v|vmware|virtualbox|virtual|vpn|tap-|tunnel|loopback|bluetooth|wan miniport|npcap|tailscale|zerotier|wireguard|docker|vethernet|fortinet|anyconnect|teredo|isatap|pangp|openvpn'

#region Normalizzazioni

function ConvertTo-HDMac {
    # "00-11-22-aa-bb-cc", "0011.22aa.bbcc", "001122AABBCC" -> "00:11:22:AA:BB:CC"
    param([string]$Testo)
    $h = ([string]$Testo).ToUpperInvariant() -replace '[^0-9A-F]', ''
    if ($h.Length -ne 12 -or $h -match '^(0{12}|F{12})$') { return '' }
    $h -replace '(..)(?!$)', '$1:'
}

function ConvertTo-HDNomeHost {
    # "pc-01.studio.local" -> "PC-01" (un indirizzo IP non è un nome)
    param([string]$Nome)
    $n = ([string]$Nome).Trim()
    if (-not $n -or $n -match $script:HDIpv4) { return '' }
    ($n -split '\.')[0].ToUpperInvariant()
}

function ConvertTo-HDProduttore {
    param([string]$Nome)
    $n = ([string]$Nome).Trim()
    switch -Regex ($n) {
        '^(hewlett[- ]?packard|hp( inc\.?)?)$' { return 'HP' }
        '^dell' { return 'Dell' }
        '^lenovo' { return 'Lenovo' }
        '^asus' { return 'ASUS' }
        '^acer' { return 'Acer' }
        '^(micro[- ]?star|msi\b)' { return 'MSI' }
        '^gigabyte' { return 'Gigabyte' }
        '^microsoft' { return 'Microsoft' }
        '^fujitsu' { return 'Fujitsu' }
        '^apple' { return 'Apple' }
        '^samsung' { return 'Samsung' }
        '^toshiba' { return 'Toshiba' }
        '^vmware' { return 'VMware' }
        '^(to be filled.*|system manufacturer|default string|o\.?e\.?m\.?)$' { return '' }
    }
    $n
}

function Get-HDModelloLeggibile {
    # Lenovo mette in Model il codice macchina (20XW00GBIX): il nome vero è in Version
    param([string]$Produttore, [string]$Modello, [string]$VersioneProdotto)
    $m = ([string]$Modello).Trim()
    $v = ([string]$VersioneProdotto).Trim()
    if ($Produttore -match '^lenovo' -and $v -and $v -notmatch '^(?i)(lenovo|none|not available|to be filled.*|default string)$') { return $v }
    if ($m -match '^(?i)(system product name|to be filled.*|default string|not applicable|none|o\.?e\.?m\.?)$') { return '' }
    $m
}

function ConvertTo-HDNomeCpu {
    # "Intel(R) Core(TM) i5-10500 CPU @ 3.10GHz" -> "Intel Core i5-10500 @ 3.10GHz"
    param([string]$Nome)
    $n = ([string]$Nome) -replace '\((R|TM|C)\)', '' -replace '\bCPU\b', '' -replace '\s+', ' '
    ($n -replace '\s+@', ' @').Trim()
}

function Get-HDTipoDispositivo {
    # Come nel censimento: le postazioni sono "PC CLIENT", i server "SERVER" (valori in config.psd1)
    param($ChassisTypes, $PCSystemType, [string]$SistemaOperativo, [string]$Client = 'PC CLIENT', [string]$Server = 'SERVER')
    if ($SistemaOperativo -match '(?i)server') { return $Server }
    foreach ($t in @($ChassisTypes)) {
        if ($null -ne $t -and [int]$t -in 17, 23, 28) { return $Server }
    }
    if ($null -ne $PCSystemType -and [int]$PCSystemType -in 4, 5, 7) { return $Server }
    $Client
}

#endregion

#region Dati del PC -> riga del censimento

function Select-HDSchedaPrincipale {
    # La scheda "vera" con cui il PC sta in rete: IPv4 valido, gateway, non virtuale, metrica più bassa
    param($Configurazioni)
    $candidate = foreach ($c in @($Configurazioni)) {
        if (-not $c) { continue }
        $ip = @(@($c.IPAddress) | Where-Object { $_ -match $script:HDIpv4 -and $_ -notlike '169.254.*' })
        if ($ip.Count -eq 0) { continue }
        [pscustomobject]@{
            Config     = $c
            IP         = $ip[0]
            ConGateway = (@(@($c.DefaultIPGateway) | Where-Object { $_ -match $script:HDIpv4 }).Count -gt 0)
            Virtuale   = ([string]$c.Description -match $script:HDSchedeVirtuali)
            Metrica    = [int]('0' + [string]$c.IPConnectionMetric)
        }
    }
    @($candidate) | Where-Object { $_ } |
        Sort-Object @{ Expression = { $_.ConGateway }; Descending = $true }, @{ Expression = { $_.Virtuale }; Descending = $false }, @{ Expression = { $_.Metrica }; Descending = $false } |
        Select-Object -First 1
}

function Get-HDElencoMac {
    # Tutti i MAC delle schede fisiche (servono a ritrovare il PC anche se la scansione ha visto il Wi-Fi)
    param($Schede, $Configurazioni)
    $macs = New-Object System.Collections.Generic.List[string]
    foreach ($s in (@($Schede) + @($Configurazioni))) {
        if (-not $s) { continue }
        $nome = [string]$s.Name
        if (-not $nome) { $nome = [string]$s.Description }
        if ($nome -match $script:HDSchedeVirtuali) { continue }
        $m = ConvertTo-HDMac $s.MACAddress
        if ($m -and -not $macs.Contains($m)) { $macs.Add($m) }
    }
    $macs.ToArray()
}

function Get-HDUtenteRilevato {
    # Utente del PC: quello collegato, altrimenti l'ultimo che ha fatto accesso, altrimenti il
    # profilo usato più di recente. Gli account dei tecnici (config) vengono saltati.
    param([string]$UtenteConsole, [string]$UltimoAccesso, $Profili, [string[]]$AccountTecnici)
    $tecnici = @($AccountTecnici | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $breve = { param($u) (([string]$u) -split '\\')[-1].Trim() }
    $eTecnico = { param($u) $b = (& $breve $u).ToLowerInvariant(); (-not $b) -or ($tecnici -contains $b) }
    if ($UtenteConsole -and -not (& $eTecnico $UtenteConsole)) { return [pscustomobject]@{ Nome = (& $breve $UtenteConsole); Fonte = 'utente collegato' } }
    if ($UltimoAccesso -and -not (& $eTecnico $UltimoAccesso)) { return [pscustomobject]@{ Nome = (& $breve $UltimoAccesso); Fonte = 'ultimo accesso' } }
    $profilo = @($Profili) | Where-Object { $_ -and -not $_.Special -and [string]$_.LocalPath -match '[\\/]Users[\\/]' } |
        ForEach-Object { [pscustomobject]@{ Nome = ([string]$_.LocalPath -replace '^.*[\\/]', ''); Uso = $_.LastUseTime } } |
        Where-Object { -not (& $eTecnico $_.Nome) -and $_.Nome -notmatch '^(?i)(default|defaultuser\d*|public|all users)$' } |
        Sort-Object Uso -Descending | Select-Object -First 1
    if ($profilo) { return [pscustomobject]@{ Nome = $profilo.Nome; Fonte = 'profilo usato più di recente' } }
    if ($UtenteConsole) { return [pscustomobject]@{ Nome = (& $breve $UtenteConsole); Fonte = 'utente collegato (account tecnico?)' } }
    [pscustomobject]@{ Nome = ''; Fonte = '' }
}

function Get-HDDominioRilevato {
    # Nome del dominio; "No" se il PC è in un gruppo di lavoro
    param($ComputerSystem, [string]$TenantEntra, [string]$SenzaDominio = 'No')
    if ($ComputerSystem -and $ComputerSystem.PartOfDomain) { return [string]$ComputerSystem.Domain }
    if ($TenantEntra) {
        if ($TenantEntra -eq '-') { return 'Entra ID' }
        return "Entra ID ($TenantEntra)"
    }
    if ($ComputerSystem) { return $SenzaDominio }
    ''
}

function Get-HDVersioneOS {
    # "24H2 (26100.2033)"
    param($SistemaOperativo, $VersioneWindows)
    $build = ''
    if ($VersioneWindows -and $VersioneWindows.CurrentBuild) {
        $build = [string]$VersioneWindows.CurrentBuild
        if ("$($VersioneWindows.UBR)" -ne '') { $build += '.' + $VersioneWindows.UBR }
    }
    elseif ($SistemaOperativo) { $build = [string]$SistemaOperativo.BuildNumber }
    $versione = ''
    if ($VersioneWindows) {
        $versione = [string]$VersioneWindows.DisplayVersion
        if (-not $versione) { $versione = [string]$VersioneWindows.ReleaseId }
    }
    if ($versione -and $build) { return "$versione ($build)" }
    if ($versione) { return $versione }
    $build
}

function ConvertFrom-HDTipoMemoria {
    # Codici SMBIOS (Windows 10 e successivi), altrimenti il vecchio MemoryType di WMI
    param($Smbios, $Wmi)
    $mappaSmbios = @{ 18 = 'DDR'; 19 = 'DDR2'; 20 = 'DDR2'; 24 = 'DDR3'; 26 = 'DDR4'; 27 = 'LPDDR'; 28 = 'LPDDR2'; 29 = 'LPDDR3'; 30 = 'LPDDR4'; 34 = 'DDR5'; 35 = 'LPDDR5' }
    $mappaWmi = @{ 20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2'; 24 = 'DDR3'; 26 = 'DDR4' }
    if ($null -ne $Smbios -and "$Smbios" -match '^\d+$' -and $mappaSmbios.ContainsKey([int]$Smbios)) { return $mappaSmbios[[int]$Smbios] }
    if ($null -ne $Wmi -and "$Wmi" -match '^\d+$' -and $mappaWmi.ContainsKey([int]$Wmi)) { return $mappaWmi[[int]$Wmi] }
    ''
}

function ConvertFrom-HDFormatoMemoria {
    param($FormFactor)
    if ($null -eq $FormFactor -or "$FormFactor" -notmatch '^\d+$') { return '' }
    switch ([int]$FormFactor) {
        8 { return 'DIMM' }
        12 { return 'SODIMM' }
        7 { return 'SIMM' }
        11 { return 'RIMM' }
        13 { return 'SRIMM' }
    }
    ''
}

function Format-HDBanchiRam {
    # "2 x 8 GB" oppure "1 x 8 GB + 1 x 4 GB" (prima i più numerosi, poi i più grandi)
    param($Moduli)
    $gruppi = @(@($Moduli) | Where-Object { $_ -and [double]$_.Capacity -gt 0 } | Group-Object { [double]$_.Capacity } |
            Sort-Object @{ Expression = { $_.Count }; Descending = $true }, @{ Expression = { [double]$_.Name }; Descending = $true })
    if ($gruppi.Count -eq 0) { return '' }
    ($gruppi | ForEach-Object {
            $c = [double]$_.Name
            if ($c -ge 1GB) { $misura = '{0} GB' -f [math]::Round($c / 1GB) } else { $misura = '{0} MB' -f [math]::Round($c / 1MB) }
            '{0} x {1}' -f $_.Count, $misura
        }) -join ' + '
}

function Get-HDMemoriaTotale {
    # "16 GB DDR4 SODIMM (2 x 8 GB)"
    param($Moduli, $ComputerSystem)
    $banchi = @(@($Moduli) | Where-Object { $_ -and [double]$_.Capacity -gt 0 })
    $totale = [double]0
    foreach ($m in $banchi) { $totale += [double]$m.Capacity }
    if ($totale -le 0 -and $ComputerSystem) { $totale = [double]$ComputerSystem.TotalPhysicalMemory }
    if ($totale -le 0) { return '' }
    if ($totale -lt 1GB) { $testo = '{0} MB' -f [math]::Round($totale / 1MB) } else { $testo = '{0} GB' -f [math]::Round($totale / 1GB) }
    $tipi = @($banchi | ForEach-Object { ConvertFrom-HDTipoMemoria -Smbios $_.SMBIOSMemoryType -Wmi $_.MemoryType } | Where-Object { $_ } | Select-Object -Unique)
    $formati = @($banchi | ForEach-Object { ConvertFrom-HDFormatoMemoria $_.FormFactor } | Where-Object { $_ } | Select-Object -Unique)
    if ($tipi.Count -gt 0) { $testo += ' ' + ($tipi -join '/') }
    if ($formati.Count -gt 0) { $testo += ' ' + ($formati -join '/') }
    if ($banchi.Count -gt 1) { $testo += ' (' + (Format-HDBanchiRam $banchi) + ')' }
    $testo
}

function Get-HDDischiCensimento {
    # Tutti i dischi fissi: "(C:) 120 GB 25% (D:) 353 GB 53%" e "(C:) 476 GB (D:) 666 GB"
    param($Volumi)
    $fissi = @(@($Volumi) | Where-Object { $_ -and [double]$_.Size -gt 0 } | Sort-Object { [string]$_.DeviceID })
    $libero = @()
    $dimensione = @()
    $minimo = $null
    foreach ($v in $fissi) {
        $pct = [int][math]::Round([double]$v.FreeSpace / [double]$v.Size * 100)
        $libero += '({0}) {1} GB {2}%' -f $v.DeviceID, [math]::Round([double]$v.FreeSpace / 1GB), $pct
        $dimensione += '({0}) {1} GB' -f $v.DeviceID, [math]::Round([double]$v.Size / 1GB)
        if ($null -eq $minimo -or $pct -lt $minimo) { $minimo = $pct }
    }
    [pscustomobject]@{ Libero = ($libero -join ' '); Dimensione = ($dimensione -join ' '); MinimoPct = $minimo }
}

function Get-HDStatoCynet {
    param($Segnali)
    $voci = @()
    $processi = @()
    if ($Segnali) {
        $processi = @(@($Segnali.Processi) | Where-Object { $_ })
        $servizi = @(@($Segnali.Servizi) | Where-Object { $_ })
        $programmi = @(@($Segnali.Programmi) | Where-Object { $_ })
        $file = @(@($Segnali.File) | Where-Object { $_ })
        if ($processi.Count -gt 0) { $voci += 'processi in esecuzione: ' + ($processi -join ', ') }
        if ($servizi.Count -gt 0) { $voci += 'servizi: ' + ($servizi -join ', ') }
        if ($programmi.Count -gt 0) { $voci += 'programmi installati: ' + ($programmi -join ', ') }
        if ($file.Count -gt 0) { $voci += 'file: ' + ($file -join ', ') }
    }
    $valore = 'No'
    if ($voci.Count -gt 0) { $valore = 'Sì' }
    [pscustomobject]@{ Valore = $valore; Dettaglio = ($voci -join ' | '); InEsecuzione = ($processi.Count -gt 0) }
}

function Get-HDSerialeMigliore {
    # Numero di serie: BIOS, poi telaio, poi prodotto. Nei PC assemblati spesso c'è solo
    # quello della scheda madre: lo usa con il prefisso "MB".
    param($Grezzo)
    foreach ($candidato in @($Grezzo.Bios.SerialNumber, $Grezzo.Telaio.SerialNumber, $Grezzo.Prodotto.IdentifyingNumber)) {
        $s = Get-HDNumeroSerie $candidato
        if ($s) { return $s }
    }
    $s = Get-HDNumeroSerie $Grezzo.SchedaMadre.SerialNumber
    if ($s) { return 'MB ' + $s }
    ''
}

function ConvertTo-HDRecordCensimento {
    # Dati grezzi raccolti sul PC -> una riga del censimento (con qualche dato in più per l'abbinamento)
    param(
        [Parameter(Mandatory)]$Grezzo,
        [string[]]$AccountTecnici = @(),
        [string]$TipoClient = 'PC CLIENT',
        [string]$TipoServer = 'SERVER',
        [string]$SenzaDominio = 'No'
    )
    $cs = $Grezzo.ComputerSystem
    $os = $Grezzo.SistemaOperativo
    $produttoreGrezzo = ''
    $modelloGrezzo = ''
    if ($cs) { $produttoreGrezzo = [string]$cs.Manufacturer; $modelloGrezzo = [string]$cs.Model }
    $versioneProdotto = ''
    if ($Grezzo.Prodotto) { $versioneProdotto = [string]$Grezzo.Prodotto.Version }
    $caption = ''
    if ($os) { $caption = ([string]$os.Caption -replace '^Microsoft\s+', '').Trim() }

    $principale = Select-HDSchedaPrincipale $Grezzo.ConfigRete
    $tuttiMac = @(Get-HDElencoMac -Schede $Grezzo.SchedeRete -Configurazioni $Grezzo.ConfigRete)
    $ip = ''
    $tipoIp = ''
    $mac = ''
    if ($principale) {
        $ip = $principale.IP
        $mac = ConvertTo-HDMac $principale.Config.MACAddress
        if ($principale.Config.DHCPEnabled) { $tipoIp = 'DHCP' } else { $tipoIp = 'Statico' }
    }
    else {
        # PC non collegato: MAC della scheda di rete cablata (o della prima scheda fisica)
        $fisiche = @(@($Grezzo.SchedeRete) | Where-Object { $_ -and (ConvertTo-HDMac $_.MACAddress) -and [string]$_.Name -notmatch $script:HDSchedeVirtuali })
        $cablata = $fisiche | Where-Object { [string]$_.Name -notmatch '(?i)wi-?fi|wireless|wlan|802\.11' } | Select-Object -First 1
        if (-not $cablata) { $cablata = $fisiche | Select-Object -First 1 }
        if ($cablata) { $mac = ConvertTo-HDMac $cablata.MACAddress }
    }
    if ($mac -and -not ($tuttiMac -contains $mac)) { $tuttiMac = @($mac) + $tuttiMac }

    $cpu = @(@($Grezzo.Processori) | Where-Object { $_ })
    $nomeCpu = ''
    if ($cpu.Count -gt 0) {
        $nomeCpu = ConvertTo-HDNomeCpu $cpu[0].Name
        if ($cpu.Count -gt 1) { $nomeCpu += " (x$($cpu.Count))" }
    }
    $utente = Get-HDUtenteRilevato -UtenteConsole $(if ($cs) { $cs.UserName }) -UltimoAccesso $Grezzo.UltimoAccesso -Profili $Grezzo.Profili -AccountTecnici $AccountTecnici
    $dischi = Get-HDDischiCensimento -Volumi $Grezzo.Volumi
    $cynet = Get-HDStatoCynet $Grezzo.Cynet
    $tipo = ''
    if ($cs) { $tipo = Get-HDTipoDispositivo -ChassisTypes $Grezzo.Chassis -PCSystemType $cs.PCSystemType -SistemaOperativo $caption -Client $TipoClient -Server $TipoServer }
    $serialeBios = ''
    if ($Grezzo.Bios) { $serialeBios = ([string]$Grezzo.Bios.SerialNumber).Trim() }
    $extra = ConvertTo-HDDettagliPC -Grezzo $Grezzo -SistemaOperativo $caption -Server ($tipo -eq $TipoServer -or $caption -match '(?i)server') -Processore ([string](@($cpu) | Select-Object -First 1).Name)

    [pscustomobject]@{
        PSTypeName        = 'HelpDesk.Censimento'
        Hostname          = ([string]$Grezzo.NomePC).ToUpperInvariant()
        IndirizzoIP       = $ip
        TipoIP            = $tipoIp
        MAC               = $mac
        DeviceType        = $tipo
        Manufacturer      = (ConvertTo-HDProduttore $produttoreGrezzo)
        Model             = (Get-HDModelloLeggibile -Produttore $produttoreGrezzo -Modello $modelloGrezzo -VersioneProdotto $versioneProdotto)
        NomeUtente        = $utente.Nome
        Utilizzatore      = ''
        Note              = ''
        Dominio           = (Get-HDDominioRilevato -ComputerSystem $cs -TenantEntra $Grezzo.Entra -SenzaDominio $SenzaDominio)
        SerialNumber      = (Get-HDSerialeMigliore $Grezzo)
        OS                = $caption
        OSVersion         = (Get-HDVersioneOS -SistemaOperativo $os -VersioneWindows $Grezzo.VersioneWindows)
        Memory            = (Get-HDMemoriaTotale -Moduli $Grezzo.Memoria -ComputerSystem $cs)
        Processor         = $nomeCpu
        FreeDiskPct       = $dischi.Libero
        DiskSize          = $dischi.Dimensione
        Cynet             = $cynet.Valore
        BitLocker         = $extra.BitLocker
        Locazione         = ''
        TuttiMAC          = $tuttiMac
        FonteNomeUtente   = $utente.Fonte
        CynetDettaglio    = $cynet.Dettaglio
        CynetInEsecuzione = $cynet.InEsecuzione
        DiscoMinimoPct    = $dischi.MinimoPct
        SerialeBios       = $serialeBios
        Dettagli          = $extra.Dettagli
        Software          = $extra.Software
        Rilevato          = $Grezzo.Raccolto
    }
}

#endregion

#region Excel del censimento: intestazioni, righe, abbinamenti

function Find-HDColonneCensimento {
    # @{ indice colonna = testo intestazione } -> @{ Chiave = indice colonna }
    param([hashtable]$Celle)
    $norm = @{}
    foreach ($col in $Celle.Keys) { $norm[[int]$col] = ConvertTo-HDChiave ([string]$Celle[$col]) }
    $mappa = @{}
    $usate = @{}
    foreach ($passo in 'esatto', 'prefisso') {
        foreach ($d in $script:HDColonneCensimento) {
            if ($mappa.ContainsKey($d.Chiave)) { continue }
            foreach ($s in $d.Sinonimi) {
                $trovate = @($norm.Keys | Where-Object {
                        -not $usate.ContainsKey($_) -and (($passo -eq 'esatto' -and $norm[$_] -eq $s) -or ($passo -eq 'prefisso' -and $norm[$_].StartsWith($s + ' ')))
                    } | Sort-Object)
                if ($trovate.Count -gt 0) { $mappa[$d.Chiave] = $trovate[0]; $usate[$trovate[0]] = $true; break }
            }
        }
    }
    $mappa
}

function ConvertTo-HDTabellaCensimento {
    # Fogli letti da Read-HDXlsx -> tabella del censimento (foglio, intestazione, colonne, righe)
    param([Parameter(Mandatory)][object[]]$Fogli, [string]$Percorso)
    foreach ($f in $Fogli) {
        $intestazione = $null
        foreach ($n in @($f.Righe.Keys | Select-Object -First 15)) {
            $mappa = Find-HDColonneCensimento $f.Righe[$n]
            if ($mappa.ContainsKey('Hostname') -or $mappa.Count -ge 4) {
                $ultimaColonna = (@($f.Righe[$n].Keys) | ForEach-Object { [int]$_ } | Measure-Object -Maximum).Maximum
                $intestazione = @{ Riga = $n; Colonne = $mappa; UltimaColonna = [int]$ultimaColonna }
                break
            }
        }
        if (-not $intestazione) { continue }
        $righe = New-Object System.Collections.Generic.List[object]
        $ultima = $intestazione.Riga
        foreach ($n in $f.Righe.Keys) {
            if ($n -le $intestazione.Riga) { continue }
            if ($n -gt $ultima) { $ultima = $n }
            $valori = @{}
            foreach ($k in $intestazione.Colonne.Keys) {
                $v = $f.Righe[$n][$intestazione.Colonne[$k]]
                if ($null -ne $v -and ([string]$v).Trim() -ne '') { $valori[$k] = ([string]$v).Trim() }
            }
            if ($valori.Count -gt 0) { $righe.Add([pscustomobject]@{ Numero = [int]$n; Valori = $valori }) }
        }
        return [pscustomobject]@{
            Percorso         = $Percorso
            Foglio           = $f.Nome
            IndiceFoglio     = $f.Indice
            ParteFoglio      = $f.Parte
            RigaIntestazione = [int]$intestazione.Riga
            Colonne          = $intestazione.Colonne
            Mancanti         = @($script:HDColonneCensimento | Where-Object { -not $_.Aggiungibile -and -not $intestazione.Colonne.ContainsKey($_.Chiave) } | ForEach-Object { $_.Titolo })
            DaAggiungere     = @($script:HDColonneCensimento | Where-Object { $_.Aggiungibile -and -not $intestazione.Colonne.ContainsKey($_.Chiave) })
            UltimaColonna    = $intestazione.UltimaColonna
            Righe            = $righe
            UltimaRiga       = [int]$ultima
        }
    }
    throw "Nell'Excel non trovo la riga di intestazione del censimento (serve almeno la colonna 'Hostname')."
}

function Find-HDRigaCensimento {
    # Cerca la riga dello stesso dispositivo: MAC, poi numero di serie, poi hostname, infine IP.
    # Il MAC non basta se numero di serie o hostname dicono il contrario: una docking station
    # o un adattatore USB di rete passa da un portatile all'altro con il suo MAC.
    param($Righe, [string[]]$Mac, [string]$Seriale, [string]$Hostname, [string]$IP, [switch]$ConsentiIP)
    $macs = @(@($Mac) | ForEach-Object { ConvertTo-HDMac $_ } | Where-Object { $_ })
    $s = Get-HDNumeroSerie $Seriale
    $h = ConvertTo-HDNomeHost $Hostname
    if ($macs.Count -gt 0) {
        foreach ($r in $Righe) {
            $m = ConvertTo-HDMac $r.Valori['MAC']
            if (-not $m -or -not ($macs -contains $m)) { continue }
            $serieRiga = Get-HDNumeroSerie $r.Valori['SerialNumber']
            $nomeRiga = ConvertTo-HDNomeHost $r.Valori['Hostname']
            if ($s -and $serieRiga) {
                if ($s -eq $serieRiga) { return [pscustomobject]@{ Riga = $r; Motivo = "stesso MAC $m" } }
                continue
            }
            if ($h -and $nomeRiga -and $h -ne $nomeRiga) { continue }
            return [pscustomobject]@{ Riga = $r; Motivo = "stesso MAC $m" }
        }
    }
    if ($s) {
        foreach ($r in $Righe) {
            if ((Get-HDNumeroSerie $r.Valori['SerialNumber']) -eq $s) { return [pscustomobject]@{ Riga = $r; Motivo = "stesso numero di serie $s" } }
        }
    }
    if ($h) {
        foreach ($r in $Righe) {
            if ((ConvertTo-HDNomeHost $r.Valori['Hostname']) -eq $h) { return [pscustomobject]@{ Riga = $r; Motivo = "stesso hostname $h" } }
        }
    }
    if ($ConsentiIP -and $IP) {
        foreach ($r in $Righe) {
            if ($r.Valori['IndirizzoIP'] -eq $IP -and -not $r.Valori['MAC'] -and -not $r.Valori['Hostname']) { return [pscustomobject]@{ Riga = $r; Motivo = "stesso IP $IP" } }
        }
    }
    $null
}

function Get-HDTitoloColonna {
    param([string]$Chiave)
    $c = $script:HDColonneCensimento | Where-Object { $_.Chiave -eq $Chiave } | Select-Object -First 1
    if ($c) { return $c.Titolo }
    $Chiave
}

function Get-HDDifferenze {
    param([hashtable]$Esistenti, [hashtable]$Nuovi)
    $diff = @{}
    foreach ($k in $Nuovi.Keys) {
        if ([string]$Esistenti[$k] -ne [string]$Nuovi[$k]) { $diff[$k] = $Nuovi[$k] }
    }
    $diff
}

function Merge-HDValoriPC {
    # Cosa scrivere nella riga per i dati raccolti sul PC. Mai valori vuoti al posto di quelli
    # esistenti; le colonne manuali solo se compilate a mano durante il censimento.
    param([hashtable]$Esistenti, [Parameter(Mandatory)]$Record, [hashtable]$Manuali)
    $nuovi = @{}
    foreach ($k in $script:HDColonneAutomatiche) {
        $v = $Record.$k
        if ($null -ne $v -and [string]$v -ne '') { $nuovi[$k] = $v }
    }
    if ($Record.IndirizzoIP) {
        $nuovi['IndirizzoIP'] = $Record.IndirizzoIP
        if ($Record.TipoIP) { $nuovi['TipoIP'] = $Record.TipoIP }
    }
    $macRiga = ConvertTo-HDMac $Esistenti['MAC']
    if ($Record.MAC -and -not ($macRiga -and @($Record.TuttiMAC) -contains $macRiga)) { $nuovi['MAC'] = $Record.MAC }
    if ($Manuali) {
        foreach ($k in 'NomeUtente', 'Utilizzatore', 'Note', 'Locazione') {
            if ($Manuali[$k]) { $nuovi[$k] = ([string]$Manuali[$k]).Trim() }
        }
    }
    Get-HDDifferenze -Esistenti $Esistenti -Nuovi $nuovi
}

function Merge-HDValoriScansione {
    # Dalla scansione: l'IP si aggiorna sempre (è la situazione attuale), il resto solo se manca.
    # Il gruppo NetBIOS non va in Dominio: non distingue un dominio da un gruppo di lavoro.
    param([hashtable]$Esistenti, [Parameter(Mandatory)]$Dispositivo)
    $nuovi = @{}
    if ($Dispositivo.IP) { $nuovi['IndirizzoIP'] = $Dispositivo.IP }
    $seVuoto = @{ Hostname = $Dispositivo.Hostname; MAC = $Dispositivo.MAC; Manufacturer = $Dispositivo.Produttore; NomeUtente = $Dispositivo.Utente }
    foreach ($k in $seVuoto.Keys) {
        if ($seVuoto[$k] -and -not $Esistenti[$k]) { $nuovi[$k] = $seVuoto[$k] }
    }
    Get-HDDifferenze -Esistenti $Esistenti -Nuovi $nuovi
}

function Add-HDRigaInTabella {
    param($Tabella, [hashtable]$Modifiche, $Riga, [hashtable]$Diff)
    foreach ($k in $Diff.Keys) { $Riga.Valori[$k] = [string]$Diff[$k] }
    if (-not $Modifiche.ContainsKey($Riga.Numero)) { $Modifiche[$Riga.Numero] = @{} }
    foreach ($k in $Diff.Keys) { $Modifiche[$Riga.Numero][$k] = $Diff[$k] }
}

function New-HDRigaTabella {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Aggiunge una riga alla tabella in memoria')]
    param($Tabella)
    $Tabella.UltimaRiga = [int]$Tabella.UltimaRiga + 1
    $riga = [pscustomobject]@{ Numero = $Tabella.UltimaRiga; Valori = @{} }
    $Tabella.Righe.Add($riga)
    $riga
}

function Add-HDRilevazioniATabella {
    # Applica in memoria le rilevazioni dei PC: restituisce modifiche (riga -> campi) ed esiti
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][object[]]$Rilevazioni)
    $modifiche = @{}
    $esiti = foreach ($ril in $Rilevazioni) {
        $rec = $ril.Dati
        $trovata = Find-HDRigaCensimento -Righe $Tabella.Righe -Mac (@($rec.TuttiMAC) + @($rec.MAC)) -Seriale $rec.SerialNumber -Hostname $rec.Hostname
        if ($trovata) { $riga = $trovata.Riga; $azione = 'Aggiornato'; $motivo = $trovata.Motivo }
        else { $riga = New-HDRigaTabella $Tabella; $azione = 'Aggiunto'; $motivo = 'non era nel censimento' }
        $diff = Merge-HDValoriPC -Esistenti $riga.Valori -Record $rec -Manuali $ril.Manuali
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $modifiche -Riga $riga -Diff $diff
        $campi = @($script:HDColonneCensimento | Where-Object { $diff.ContainsKey($_.Chiave) } | ForEach-Object { $_.Titolo })
        [pscustomobject]@{ Hostname = $rec.Hostname; Riga = $riga.Numero; Azione = $azione; Motivo = $motivo; Campi = $campi }
    }
    [pscustomobject]@{ Modifiche = $modifiche; Esiti = @($esiti) }
}

function Add-HDScansioneATabella {
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][object[]]$Dispositivi)
    $modifiche = @{}
    $esiti = foreach ($d in $Dispositivi) {
        $trovata = Find-HDRigaCensimento -Righe $Tabella.Righe -Mac @($d.MAC) -Hostname $d.Hostname -IP $d.IP -ConsentiIP
        if ($trovata) { $riga = $trovata.Riga; $azione = 'Aggiornato' }
        else { $riga = New-HDRigaTabella $Tabella; $azione = 'Aggiunto' }
        $diff = Merge-HDValoriScansione -Esistenti $riga.Valori -Dispositivo $d
        if ($azione -eq 'Aggiornato' -and $diff.Count -eq 0) { $azione = 'Invariato' }
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $modifiche -Riga $riga -Diff $diff
        [pscustomobject]@{ Hostname = $d.Hostname; IP = $d.IP; Riga = $riga.Numero; Azione = $azione }
    }
    [pscustomobject]@{ Modifiche = $modifiche; Esiti = @($esiti) }
}

function ConvertTo-HDModificheXlsx {
    # Modifiche per chiave di colonna -> modifiche per indice di colonna (solo colonne presenti nell'Excel)
    param([Parameter(Mandatory)]$Tabella, [Parameter(Mandatory)][hashtable]$Modifiche)
    $lista = foreach ($n in @($Modifiche.Keys | Sort-Object)) {
        $celle = @{}
        foreach ($k in $Modifiche[$n].Keys) {
            if ($Tabella.Colonne.ContainsKey($k)) { $celle[[int]$Tabella.Colonne[$k]] = ([string]$Modifiche[$n][$k]).Trim() }
        }
        if ($celle.Count -gt 0) { @{ Riga = [int]$n; Celle = $celle } }
    }
    $lista
}

#endregion

#region Scansione di rete (Advanced IP Scanner: export CSV o XML)

$script:HDCampiScansione = [ordered]@{
    Stato      = @('status', 'stato', 'state')
    Nome       = @('name', 'nome', 'hostname', 'hostnome', 'nomehost', 'computer', 'nomecomputer', 'dnsname', 'nomedns', 'netbiosname', 'nomenetbios')
    IP         = @('ip', 'ipaddress', 'indirizzoip', 'indirizzo', 'ipv4', 'ipv4address')
    Produttore = @('manufacturer', 'produttore', 'vendor', 'fornitore', 'costruttore', 'nicvendor', 'macvendor')
    MAC        = @('macaddress', 'mac', 'indirizzomac', 'macaddr', 'physicaladdress', 'indirizzofisico')
    Commenti   = @('comments', 'comment', 'commenti', 'commento')
    Utente     = @('user', 'utente', 'nomeutente', 'username', 'loggeduser')
    Gruppo     = @('netbiosgroup', 'gruppo', 'grupponetbios', 'workgroup', 'gruppodilavoro', 'domain', 'dominio', 'group')
}

function ConvertTo-HDChiaveCompatta {
    param([string]$Testo)
    (ConvertTo-HDChiave $Testo) -replace '[ \-]', ''
}

function Get-HDMappaScansione {
    # nomi delle colonne -> @{ campo = posizione }
    param([string[]]$Nomi)
    $compatti = @($Nomi | ForEach-Object { ConvertTo-HDChiaveCompatta $_ })
    $mappa = @{}
    foreach ($campo in $script:HDCampiScansione.Keys) {
        foreach ($s in $script:HDCampiScansione[$campo]) {
            $i = [array]::IndexOf($compatti, $s)
            if ($i -ge 0 -and -not ($mappa.Values -contains $i)) { $mappa[$campo] = $i; break }
        }
    }
    $mappa
}

function ConvertFrom-HDTestoCsv {
    # CSV con virgolette e a capo dentro i campi; separatore a scelta
    param([AllowEmptyString()][string]$Testo, [char]$Separatore)
    $righe = New-Object System.Collections.Generic.List[object]
    $campi = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $virgolette = $false
    $n = $Testo.Length
    for ($i = 0; $i -lt $n; $i++) {
        $ch = $Testo[$i]
        if ($virgolette) {
            if ($ch -eq [char]'"') {
                if ($i + 1 -lt $n -and $Testo[$i + 1] -eq [char]'"') { [void]$sb.Append('"'); $i++ }
                else { $virgolette = $false }
            }
            else { [void]$sb.Append($ch) }
        }
        elseif ($ch -eq [char]'"') { $virgolette = $true }
        elseif ($ch -eq $Separatore) { $campi.Add($sb.ToString()); [void]$sb.Clear() }
        elseif ($ch -eq [char]"`n") { $campi.Add($sb.ToString()); [void]$sb.Clear(); $righe.Add($campi.ToArray()); $campi.Clear() }
        elseif ($ch -ne [char]"`r") { [void]$sb.Append($ch) }
    }
    if ($sb.Length -gt 0 -or $campi.Count -gt 0) { $campi.Add($sb.ToString()); $righe.Add($campi.ToArray()) }
    , $righe
}

function New-HDDispositivoScansione {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Crea solo un oggetto in memoria')]
    param([hashtable]$Valori)
    $testo = { param($k) ([string]$Valori[$k]).Trim() }
    $ip = ''
    if ((& $testo 'IP') -match '(\d{1,3}(\.\d{1,3}){3})') { $ip = $Matches[1] }
    [pscustomobject]@{
        Hostname   = (ConvertTo-HDNomeHost (& $testo 'Nome'))
        Nome       = (& $testo 'Nome')
        IP         = $ip
        MAC        = (ConvertTo-HDMac (& $testo 'MAC'))
        Produttore = (& $testo 'Produttore')
        Stato      = (& $testo 'Stato')
        Utente     = ((& $testo 'Utente') -split '\\')[-1]
        Gruppo     = (& $testo 'Gruppo')
        Commenti   = (& $testo 'Commenti')
    }
}

function Test-HDDispositivoAttivo {
    param($Dispositivo, [bool]$ColonnaStato)
    if (-not $Dispositivo.IP) { return $false }
    if ($ColonnaStato -and $Dispositivo.Stato) {
        return -not ((ConvertTo-HDChiave $Dispositivo.Stato) -match '^(off|offline|dead|morto|spento|inattivo|non attivo|non in linea|non raggiungibile|disattivo|down|unknown|sconosciuto)')
    }
    [bool]($Dispositivo.Hostname -or $Dispositivo.MAC)
}

function ConvertFrom-HDScansione {
    <#
    Legge l'export di Advanced IP Scanner (CSV con ; , o tab, oppure XML; intestazioni in
    italiano o in inglese) e restituisce i dispositivi attivi, senza doppioni.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Testo)
    $grezzi = New-Object System.Collections.Generic.List[object]
    $formato = ''
    $colonne = @()
    if ($Testo.TrimStart().StartsWith('<')) {
        $formato = 'XML'
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        try { $doc.LoadXml($Testo) } catch { throw "Il file XML della scansione non è leggibile: $($_.Exception.Message)" }
        foreach ($el in $doc.SelectNodes('//*')) {
            $nomi = New-Object System.Collections.Generic.List[string]
            $valori = New-Object System.Collections.Generic.List[string]
            foreach ($a in $el.Attributes) { $nomi.Add($a.LocalName); $valori.Add($a.Value) }
            foreach ($figlio in $el.ChildNodes) {
                if ($figlio.NodeType -eq [System.Xml.XmlNodeType]::Element -and -not $figlio.SelectSingleNode('*')) { $nomi.Add($figlio.LocalName); $valori.Add($figlio.InnerText) }
            }
            $mappa = Get-HDMappaScansione -Nomi $nomi.ToArray()
            if (-not $mappa.ContainsKey('IP')) { continue }
            $v = @{}
            foreach ($campo in $mappa.Keys) { $v[$campo] = $valori[$mappa[$campo]] }
            $grezzi.Add(@{ Valori = $v; ConStato = $mappa.ContainsKey('Stato') })
            if (-not $colonne) { $colonne = @($mappa.Keys) }
        }
    }
    else {
        $migliore = $null
        foreach ($sep in [char[]]@(';', ',', "`t")) {
            $righe = ConvertFrom-HDTestoCsv -Testo $Testo -Separatore $sep
            for ($i = 0; $i -lt [math]::Min(10, $righe.Count); $i++) {
                $mappa = Get-HDMappaScansione -Nomi $righe[$i]
                if (-not $mappa.ContainsKey('IP') -or $mappa.Count -lt 2) { continue }
                if (-not $migliore -or $mappa.Count -gt $migliore.Mappa.Count) { $migliore = @{ Righe = $righe; Intestazione = $i; Mappa = $mappa; Separatore = $sep } }
                break
            }
        }
        if (-not $migliore) {
            $prima = ''
            if ($Testo) { $prima = ($Testo -split "\r?\n")[0] }
            throw "Non riconosco le colonne della scansione (servono almeno IP e nome o MAC). Prima riga del file: '$prima'"
        }
        $formato = 'CSV (separatore ' + $(if ($migliore.Separatore -eq "`t") { 'tabulazione' } else { "'$($migliore.Separatore)'" }) + ')'
        $colonne = @($migliore.Mappa.Keys)
        for ($i = $migliore.Intestazione + 1; $i -lt $migliore.Righe.Count; $i++) {
            $campi = $migliore.Righe[$i]
            if (-not ($campi -join '').Trim()) { continue }
            $v = @{}
            foreach ($campo in $migliore.Mappa.Keys) {
                $pos = $migliore.Mappa[$campo]
                if ($pos -lt $campi.Count) { $v[$campo] = $campi[$pos] }
            }
            $grezzi.Add(@{ Valori = $v; ConStato = $migliore.Mappa.ContainsKey('Stato') })
        }
    }

    $visti = @{}
    $dispositivi = New-Object System.Collections.Generic.List[object]
    $ignorati = 0
    foreach ($g in $grezzi) {
        $d = New-HDDispositivoScansione $g.Valori
        if (-not (Test-HDDispositivoAttivo -Dispositivo $d -ColonnaStato $g.ConStato)) { $ignorati++; continue }
        $chiave = $d.MAC
        if (-not $chiave) { $chiave = $d.IP }
        if ($visti.ContainsKey($chiave)) { continue }
        $visti[$chiave] = $true
        $dispositivi.Add($d)
    }
    [pscustomobject]@{ Dispositivi = $dispositivi.ToArray(); Ignorati = $ignorati; Formato = $formato; Colonne = $colonne }
}

#endregion

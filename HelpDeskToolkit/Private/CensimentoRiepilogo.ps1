# Censimento: riepilogo dei problemi del parco PC (foglio "Riepilogo" dell'Excel, una riga per PC, ricalcolato
# a ogni scrittura, e report HTML da consegnare) e storico dei cambiamenti tra un censimento e l'altro
# (foglio "Storico", una riga per cambiamento). Funzioni pure: lavorano sulle tabelle lette dall'Excel.

$script:HDFoglioRiepilogo = 'Riepilogo'
$script:HDFoglioStorico = 'Storico'

# Foglio Riepilogo: una riga per PC con i suoi problemi (riscritto ogni volta, non va modificato a mano)
$script:HDColonneRiepilogo = @(
    New-HDColonnaFoglio 'Hostname' 'Hostname' 20
    New-HDColonnaFoglio 'Utilizzatore' 'Utilizzatore' 22
    New-HDColonnaFoglio 'Urgente' 'Urgente' 60
    New-HDColonnaFoglio 'DaControllare' 'Da controllare' 90
)

# Foglio Storico: una riga per cambiamento (una frase), si aggiunge in fondo a ogni censimento
$script:HDColonneStorico = @(
    New-HDColonnaFoglio 'Data' 'Data' 17
    New-HDColonnaFoglio 'Hostname' 'Hostname' 20
    New-HDColonnaFoglio 'SerialNumber' 'Serial Number' 18
    New-HDColonnaFoglio 'Cambiamento' 'Cosa è cambiato' 100
)

# Fine degli aggiornamenti di sicurezza di Windows 11 per versione: Home/Pro, Enterprise/Education
$script:HDFineSupportoWindows11 = @{
    '21H2' = @('2023-10-10', '2024-10-08')
    '22H2' = @('2024-10-08', '2025-10-14')
    '23H2' = @('2025-11-11', '2026-11-10')
    '24H2' = @('2026-10-13', '2027-10-12')
    '25H2' = @('2027-10-12', '2028-10-10')
}
$script:HDFineSupportoWindows10 = '2025-10-14'

# Colonne seguite nello storico (foglio principale e Dettagli). Le altre cambiano troppo spesso
# (IP, spazio libero, stampanti, data degli aggiornamenti) o non dicono niente di nuovo.
$script:HDCampiStorico = @('Hostname', 'SerialNumber', 'Manufacturer', 'Model', 'NomeUtente', 'Dominio', 'OS', 'OSVersion', 'Memory', 'Processor', 'DiskSize', 'Cynet', 'BitLocker')
$script:HDCampiStoricoDettagli = @('DischiFisici', 'Monitor', 'Bios', 'Office', 'AttivazioneWindows', 'Antivirus', 'AccessoRemoto', 'Amministratori')

# Nomi dei campi nelle frasi dello storico, e campi che sono elenchi (con il loro separatore)
$script:HDEtichetteStorico = @{
    Hostname = 'Nome del PC'; SerialNumber = 'Numero di serie'; Manufacturer = 'Produttore'; Model = 'Modello'
    NomeUtente = 'Utente'; Dominio = 'Dominio'; OS = 'Sistema operativo'; OSVersion = 'Versione di Windows'
    Memory = 'RAM'; Processor = 'Processore'; DiskSize = 'Dischi'; Cynet = 'Cynet'; BitLocker = 'BitLocker'
    DischiFisici = 'Dischi fisici'; Monitor = 'Monitor'; Bios = 'BIOS'; Office = 'Office'
    AttivazioneWindows = 'Attivazione di Windows'; Antivirus = 'Antivirus'; AccessoRemoto = 'Accesso remoto'
    Amministratori = 'Amministratori'
}
$script:HDCampiElencoStorico = @{ BitLocker = ';'; DischiFisici = ';'; Monitor = ';'; Office = ';'; Antivirus = ';'; AccessoRemoto = ';'; Amministratori = ',' }

#region Riepilogo

function Get-HDSoglieRiepilogo {
    # Soglie del riepilogo: quelle della configurazione, con i valori predefiniti per quelle che mancano
    param([hashtable]$Soglie)
    $base = @{
        DiscoAttenzionePct = 10; DiscoAttenzioneGB = 20; DiscoCriticoPct = 5; DiscoCriticoGB = 10
        AggiornamentiGiorniAttenzione = 45; RamMinimaGB = 8; BiosAnni = 6; CensimentoMesi = 6
    }
    if ($Soglie) { foreach ($k in @($base.Keys)) { if ($null -ne $Soglie[$k] -and "$($Soglie[$k])" -ne '') { $base[$k] = [double]$Soglie[$k] } } }
    $base
}

function ConvertFrom-HDDataTesto {
    # Prima data gg/mm/aaaa di un testo ("14/10/2026 (KB5066835)", "1.15.0 (12/03/2022)") -> [datetime] o $null
    param([string]$Testo)
    if ([string]$Testo -notmatch '(\d{1,2})/(\d{1,2})/(\d{4})') { return $null }
    try { New-Object datetime ([int]$Matches[3]), ([int]$Matches[2]), ([int]$Matches[1]) }
    catch { $null }
}

function Get-HDVolumiTesto {
    # "(C:) 120 GB 25% (D:) 353 GB 53%" -> Unita, LiberoGB, Pct
    param([string]$Testo)
    foreach ($m in [regex]::Matches([string]$Testo, '\((\w:)\)\s*(\d+(?:[.,]\d+)?)\s*GB\s*(\d+)\s*%')) {
        [pscustomobject]@{ Unita = $m.Groups[1].Value; LiberoGB = [double]($m.Groups[2].Value -replace ',', '.'); Pct = [int]$m.Groups[3].Value }
    }
}

function Get-HDMemoriaGB {
    # "16 GB DDR4 SODIMM (2 x 8 GB)" -> 16; "512 MB" -> 0,5; testo vuoto o diverso -> 0
    param([string]$Testo)
    if ([string]$Testo -notmatch '^\s*(\d+(?:[.,]\d+)?)\s*(GB|MB)') { return 0 }
    $n = [double]($Matches[1] -replace ',', '.')
    if ($Matches[2] -eq 'MB') { return $n / 1024 }
    $n
}

function Get-HDWindowsSenzaAggiornamenti {
    # Il Windows del PC non riceve più aggiornamenti di sicurezza? '' se è ancora aggiornato (o se non si sa),
    # altrimenti il nome breve ("Windows 10", "Windows 11 23H2", "Windows 7")
    param([string]$SistemaOperativo, [string]$Versione, [datetime]$Adesso = (Get-Date))
    $os = ([string]$SistemaOperativo).Trim()
    if ($os -match '(?i)windows (xp|vista|7|8\.1|8)\b') { return 'Windows ' + $Matches[1] }
    if ($os -match '(?i)server (2003|2008|2012)') { return 'Windows Server ' + $Matches[1] }
    if ($os -match '(?i)windows 10') {
        if ($os -match '(?i)LTS[BC]|server') { return '' }
        if ($Adesso -ge [datetime]::ParseExact($script:HDFineSupportoWindows10, 'yyyy-MM-dd', $script:HDInvariante)) { return 'Windows 10' }
        return ''
    }
    if ($os -match '(?i)windows 11' -and [string]$Versione -match '^\s*(\d{2}H\d)') {
        $ver = $Matches[1].ToUpperInvariant()
        $date = $script:HDFineSupportoWindows11[$ver]
        if (-not $date) { return '' }
        $i = 0
        if ($os -match '(?i)enterprise|education') { $i = 1 }
        if ($Adesso -ge [datetime]::ParseExact($date[$i], 'yyyy-MM-dd', $script:HDInvariante)) { return 'Windows 11 ' + $ver }
    }
    ''
}

function Get-HDPcCensiti {
    # PC del foglio principale con i loro dati del foglio Dettagli (stesso numero di serie, altrimenti stesso
    # nome). Solo i PC censiti (sistema operativo o dettagli): i dispositivi visti solo dalla scansione no.
    param([object[]]$Righe, [object[]]$Dettagli)
    $perSerie = @{}
    $perNome = @{}
    foreach ($d in @($Dettagli | Where-Object { $_ })) {
        $s = Get-HDNumeroSerie $d.Valori['SerialNumber']
        if ($s -and -not $perSerie.ContainsKey($s)) { $perSerie[$s] = $d.Valori }
        $h = ConvertTo-HDNomeHost $d.Valori['Hostname']
        if ($h -and -not $perNome.ContainsKey($h)) { $perNome[$h] = $d.Valori }
    }
    foreach ($r in @($Righe | Where-Object { $_ })) {
        $v = $r.Valori
        $s = Get-HDNumeroSerie $v['SerialNumber']
        $h = ConvertTo-HDNomeHost $v['Hostname']
        $det = $null
        if ($s -and $perSerie.ContainsKey($s)) { $det = $perSerie[$s] }
        elseif ($h -and $perNome.ContainsKey($h)) { $det = $perNome[$h] }
        if (-not $v['OS'] -and -not $det) { continue }
        if (-not $det) { $det = @{} }
        $nome = [string]$v['Hostname']
        if (-not $nome) { $nome = [string]$v['SerialNumber'] }
        if (-not $nome) { $nome = [string]$v['IndirizzoIP'] }
        [pscustomobject]@{ Nome = $nome; Utilizzatore = [string]$v['Utilizzatore']; Valori = $v; Dettagli = $det }
    }
}

function Join-HDElenco {
    # @('A', 'B', 'C') -> "A, B e C"
    param([string[]]$Voci)
    $v = @($Voci | Where-Object { $_ })
    if ($v.Count -le 1) { return ($v -join '') }
    (($v | Select-Object -First ($v.Count - 1)) -join ', ') + ' e ' + $v[-1]
}

function Get-HDDischiARischio {
    # Nomi dei dischi segnalati da SMART o da Windows ("ATTENZIONE", "GUASTO"); "" se c'è solo la previsione
    # di guasto dello SMART senza nome. $null se nessun disco è a rischio.
    param([string]$Smart, [string]$DischiFisici)
    $nomi = New-Object System.Collections.Generic.List[string]
    $trovato = $false
    foreach ($voce in @(([string]$Smart -split ';\s*') + ([string]$DischiFisici -split ';\s*'))) {
        if ($voce -cnotmatch 'GUASTO|ATTENZIONE') { continue }
        $trovato = $true
        if ($voce -cmatch '^GUASTO PREVISTO') { continue }
        $nome = ($voce -replace '(:\s.*|\s\(.*)$', '').Trim()
        if ($nome -and -not $nomi.Contains($nome)) { $nomi.Add($nome) }
    }
    if (-not $trovato) { return $null }
    Join-HDElenco $nomi.ToArray()
}

function Get-HDPuntiRiepilogo {
    # Problemi del parco PC, controllo per controllo: Priorita (Urgente, Da controllare), Punto (nome breve,
    # per contare i PC) e PC (Nome, Utilizzatore, Fatto: la frase precisa per quel PC, es. "BitLocker non
    # attivo (portatile)"). Solo i controlli che riguardano almeno un PC.
    param([object[]]$PC, [hashtable]$Soglie, [datetime]$Adesso = (Get-Date))
    $s = Get-HDSoglieRiepilogo $Soglie
    $usaCynet = @($PC | Where-Object { $_ -and [string]$_.Valori['Cynet'] -match '^S' }).Count -gt 0
    $quando = {
        # data del censimento del PC (per confrontare le date di quel momento), altrimenti oggi
        param($p)
        $d = ConvertFrom-HDDataTesto $p.Dettagli['Rilevato']
        if ($d) { return $d }
        $Adesso
    }
    $controlli = @(
        @{
            Priorita = 'Urgente'; Punto = 'Disco a rischio di guasto'
            Prova    = {
                param($p)
                $nomi = Get-HDDischiARischio -Smart $p.Dettagli['Smart'] -DischiFisici $p.Dettagli['DischiFisici']
                if ($null -eq $nomi) { return $false }
                if (-not $nomi) { return 'Disco a rischio di guasto (lo segnala lo SMART)' }
                'Disco a rischio di guasto: ' + $nomi
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Windows senza aggiornamenti di sicurezza'
            Prova    = {
                param($p)
                $vecchio = Get-HDWindowsSenzaAggiornamenti -SistemaOperativo $p.Valori['OS'] -Versione $p.Valori['OSVersion'] -Adesso $Adesso
                if (-not $vecchio) { return $false }
                $fatto = $vecchio + ' senza aggiornamenti di sicurezza'
                if ($vecchio -like 'Windows 11*') { return $fatto }
                $w11 = [string]$p.Dettagli['Windows11']
                if ($w11 -match '^S') { return $fatto + ' (compatibile con Windows 11)' }
                if ($w11 -match '^No') { return $fatto + ' (non compatibile con Windows 11)' }
                $fatto
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Antivirus spento'
            Prova    = {
                param($p)
                $av = [string]$p.Dettagli['Antivirus']
                if (-not $av -or $av -match '\(attivo') { return $false }
                $nomi = @($av -split ';\s*' | ForEach-Object { ($_ -replace '\s\(.*$', '').Trim() } | Where-Object { $_ })
                'Antivirus spento: ' + (Join-HDElenco $nomi)
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Cynet non installato'
            # solo dove Cynet è in uso (almeno un PC lo ha): per gli altri clienti non è un problema
            Prova    = { param($p) if ($usaCynet -and [string]$p.Valori['Cynet'] -eq 'No') { 'Cynet non installato' } else { $false } }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Disco pieno'
            Prova    = {
                param($p)
                $pieni = @(Get-HDVolumiTesto $p.Valori['FreeDiskPct'] | Where-Object { $_.Pct -lt $s.DiscoCriticoPct -or $_.LiberoGB -lt $s.DiscoCriticoGB })
                if ($pieni.Count -eq 0) { return $false }
                (@($pieni | ForEach-Object { 'Disco {0} pieno ({1} GB liberi)' -f $_.Unita, $_.LiberoGB }) -join '; ')
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Windows non attivato'
            Prova    = { param($p) if ([string]$p.Dettagli['AttivazioneWindows'] -like 'Non attivato*') { 'Windows non attivato' } else { $false } }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Aggiornamenti di Windows fermi'
            Prova    = {
                param($p)
                $ultimo = ConvertFrom-HDDataTesto $p.Dettagli['Aggiornamento']
                if (-not $ultimo) { return $false }
                if (((& $quando $p) - $ultimo).TotalDays -le $s.AggiornamentiGiorniAttenzione) { return $false }
                'Aggiornamenti di Windows fermi dal ' + $ultimo.ToString('dd/MM/yyyy', $script:HDInvariante)
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Disco quasi pieno'
            Prova    = {
                param($p)
                $volumi = @(Get-HDVolumiTesto $p.Valori['FreeDiskPct'])
                $pochi = @($volumi | Where-Object { -not ($_.Pct -lt $s.DiscoCriticoPct -or $_.LiberoGB -lt $s.DiscoCriticoGB) -and ($_.Pct -lt $s.DiscoAttenzionePct -or $_.LiberoGB -lt $s.DiscoAttenzioneGB) })
                if ($pochi.Count -eq 0) { return $false }
                (@($pochi | ForEach-Object { 'Disco {0} quasi pieno ({1} GB liberi)' -f $_.Unita, $_.LiberoGB }) -join '; ')
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'BitLocker non attivo'
            Prova    = {
                param($p)
                $b = [string]$p.Valori['BitLocker']
                if (-not $b) { return $false }
                $sistema = @($b -split ';' | Where-Object { $_.Trim() -match '^C:' }) | Select-Object -First 1
                if (-not $sistema) { $sistema = @($b -split ';')[0] }
                # "non attiva", "sospesa", "decrittografia in corso", "in attesa di attivazione": dati non protetti
                if ($sistema.Trim() -notmatch '^\w:\s*(non attiva|sospesa|decrittografia|in attesa)') { return $false }
                $batteria = [string]$p.Dettagli['Batteria']
                if ($batteria -and $batteria -ne 'Nessuna') { return 'BitLocker non attivo (portatile)' }
                'BitLocker non attivo'
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Utente amministratore del PC'
            Prova    = {
                param($p)
                $a = [string]$p.Dettagli['Amministratori']
                if ($a -notmatch '(?:^|,\s*)([^,(]+?)\s\((?:[^)]*,\s*)?utente del PC') { return $false }
                "L'utente {0} è amministratore" -f $Matches[1].Trim()
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Desktop remoto (RDP) attivo'
            Prova    = { param($p) if ([string]$p.Dettagli['AccessoRemoto'] -match 'Desktop remoto') { 'Desktop remoto (RDP) attivo' } else { $false } }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Programmi di accesso remoto'
            Prova    = {
                param($p)
                $programmi = @(([string]$p.Dettagli['AccessoRemoto'] -split ';\s*') | Where-Object { $_ -and $_ -ne 'Nessuno' -and $_ -notmatch 'Desktop remoto' })
                if ($programmi.Count -eq 0) { return $false }
                if ($programmi.Count -eq 1) { return $programmi[0] + ' installato' }
                (Join-HDElenco $programmi) + ' installati'
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Office non più supportato'
            Prova    = {
                param($p)
                $o = [string]$p.Dettagli['Office']
                if ($o -match '(?<!\d)(2007|2010|2013|2016|2019|2021)(?!\d)') { return ('Office {0} non più supportato' -f $Matches[1]) }
                $false
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Batteria usurata'
            Prova    = {
                param($p)
                $b = [string]$p.Dettagli['Batteria']
                if ($b -match '^(\d+)%.*da sostituire') { return ('Batteria al {0}%, da sostituire' -f $Matches[1]) }
                if ($b -match '^(\d+)%.*usurata') { return ('Batteria al {0}%' -f $Matches[1]) }
                $false
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Poca RAM'
            Prova    = {
                param($p)
                $gb = Get-HDMemoriaGB $p.Valori['Memory']
                if ($gb -gt 0 -and $gb -lt ($s.RamMinimaGB - 0.5)) { return ('Solo {0} GB di RAM' -f [math]::Round($gb, 1)) }
                $false
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'BIOS vecchio'
            Prova    = {
                param($p)
                $d = ConvertFrom-HDDataTesto $p.Dettagli['Bios']
                if (-not $d) { return $false }
                if (((& $quando $p) - $d).TotalDays -lt 365.25 * $s.BiosAnni) { return $false }
                'BIOS del ' + $d.Year
            }
        }
        @{
            Priorita = 'Da controllare'; Punto = 'Non censito da tempo'
            Prova    = {
                param($p)
                $d = ConvertFrom-HDDataTesto $p.Dettagli['Rilevato']
                if (-not $d -or $d -gt $Adesso.AddMonths(-1 * [int]$s.CensimentoMesi)) { return $false }
                'Non censito dal ' + $d.ToString('dd/MM/yyyy', $script:HDInvariante)
            }
        }
    )
    foreach ($c in $controlli) {
        $colpiti = New-Object System.Collections.Generic.List[object]
        foreach ($p in @($PC | Where-Object { $_ })) {
            $fatto = & $c.Prova $p
            if (-not $fatto) { continue }
            $colpiti.Add([pscustomobject]@{ Nome = $p.Nome; Utilizzatore = $p.Utilizzatore; Fatto = [string]$fatto })
        }
        if ($colpiti.Count -gt 0) {
            [pscustomobject]@{ Priorita = $c.Priorita; Punto = $c.Punto; PC = $colpiti.ToArray() }
        }
    }
}

function Get-HDRiepilogoPerPc {
    # Un elemento per PC con i suoi problemi (Urgente e DaControllare: le frasi), prima i PC messi peggio
    param([object[]]$PC, [object[]]$Punti)
    $perPc = [ordered]@{}
    foreach ($p in @($PC | Where-Object { $_ })) {
        if ($perPc.Contains($p.Nome)) { continue }
        $perPc[$p.Nome] = [pscustomobject]@{ Nome = $p.Nome; Utilizzatore = $p.Utilizzatore; Urgente = (New-Object System.Collections.Generic.List[string]); DaControllare = (New-Object System.Collections.Generic.List[string]) }
    }
    foreach ($punto in @($Punti | Where-Object { $_ })) {
        foreach ($x in @($punto.PC)) {
            if (-not $perPc.Contains($x.Nome)) { continue }
            if ($punto.Priorita -eq 'Urgente') { $perPc[$x.Nome].Urgente.Add($x.Fatto) } else { $perPc[$x.Nome].DaControllare.Add($x.Fatto) }
        }
    }
    $perPc.Values | Sort-Object @{ Expression = { $_.Urgente.Count }; Descending = $true }, @{ Expression = { $_.DaControllare.Count }; Descending = $true }, Nome
}

function ConvertTo-HDRigheRiepilogo {
    # PC con i loro problemi -> righe del foglio Riepilogo (una per PC), poi la riga con la data
    param([object[]]$PerPc, [datetime]$Adesso = (Get-Date))
    $elenco = @($PerPc | Where-Object { $_ })
    foreach ($p in $elenco) {
        , @($p.Nome, $p.Utilizzatore, ($p.Urgente.ToArray() -join '; '), ($p.DaControllare.ToArray() -join '; '))
    }
    $nota = 'Calcolato il {0} su {1}: si aggiorna da solo a ogni censimento, non scriverci dentro.' -f (Format-HDData $Adesso), (Format-HDQuantita $elenco.Count 'PC censito' 'PC censiti')
    , @($nota, '', '', '')
}

function Get-HDRigheRiepilogoTabella {
    # Righe del foglio Riepilogo dai dati (già aggiornati in memoria) del foglio principale e di Dettagli
    param([Parameter(Mandatory)]$Tabella, [hashtable]$Soglie, [datetime]$Adesso = (Get-Date))
    $dettagli = @()
    if ($Tabella.Dettagli) { $dettagli = @($Tabella.Dettagli.Righe | Where-Object { $_ }) }
    $pc = @(Get-HDPcCensiti -Righe @($Tabella.Righe | Where-Object { $_ }) -Dettagli $dettagli)
    $punti = @(Get-HDPuntiRiepilogo -PC $pc -Soglie $Soglie -Adesso $Adesso)
    ConvertTo-HDRigheRiepilogo -PerPc @(Get-HDRiepilogoPerPc -PC $pc -Punti $punti) -Adesso $Adesso
}

function Get-HDStatoFoglioRiscritto {
    # Fogli che il toolkit riscrive ogni volta (Software, Riepilogo): vanno bene se mancano, se sono vuoti
    # o se sono nostri (la prima riga contiene il titolo indicato); con altro contenuto non si toccano
    param([object[]]$Fogli, [Parameter(Mandatory)][string]$Nome, [Parameter(Mandatory)][object[]]$Colonne, [Parameter(Mandatory)][string]$Titolo)
    $f = @($Fogli | Where-Object { $_ -and $_.Nome -eq $Nome }) | Select-Object -First 1
    $stato = [pscustomobject]@{ Nome = $Nome; Definizione = $Colonne; Esiste = $false; Conflitto = $false; ParteFoglio = $null; IndiceFoglio = -1; Nascosto = $false }
    if (-not $f) { return $stato }
    $stato.Esiste = $true
    $stato.ParteFoglio = $f.Parte
    $stato.IndiceFoglio = $f.Indice
    if ($f.Righe.Count -gt 0) {
        $primaRiga = $f.Righe[@($f.Righe.Keys)[0]]
        $titoli = @($primaRiga.Values | ForEach-Object { ConvertTo-HDChiave ([string]$_) })
        if ($titoli -notcontains (ConvertTo-HDChiave $Titolo)) { $stato.Conflitto = $true }
    }
    $stato
}

function Get-HDNomeCliente {
    # "Censimento_Studio-Rossi.xlsx" -> "Studio Rossi"; "Censimento.xlsx" -> ""
    param([string]$Percorso)
    $nome = ([string]$Percorso -split '[\\/]')[-1] -replace '\.xlsx$', ''
    $nome = $nome -replace '^(?i)censimento[ _-]*', '' -replace '[ _-]*\(\d+\)$', ''
    ($nome -replace '[_-]', ' ' -replace '\s+', ' ').Trim()
}

#endregion

#region Storico

function Get-HDFotoCensimento {
    # Copia dei valori di tutte le righe (foglio principale, Dettagli, DatiSoftware) prima di applicare i
    # censimenti: dopo si confronta per sapere cosa è cambiato
    param([Parameter(Mandatory)]$Tabella)
    $copia = {
        param($Righe)
        $foto = @{}
        foreach ($r in @($Righe | Where-Object { $_ })) {
            $valori = @{}
            foreach ($k in $r.Valori.Keys) { $valori[$k] = $r.Valori[$k] }
            $foto[[int]$r.Numero] = $valori
        }
        $foto
    }
    $foto = @{ Principale = (& $copia $Tabella.Righe); Dettagli = @{}; Software = @{} }
    if ($Tabella.Dettagli) { $foto.Dettagli = & $copia $Tabella.Dettagli.Righe }
    if ($Tabella.DatiSoftware) { $foto.Software = & $copia $Tabella.DatiSoftware.Righe }
    $foto
}

function ConvertTo-HDValoreStorico {
    # Valore da confrontare: senza le parti che cambiano a ogni aggiornamento (build di Windows, versione
    # di Office a consumo, stato delle firme dell'antivirus)
    param([string]$Campo, [string]$Valore)
    $v = ([string]$Valore).Trim()
    switch ($Campo) {
        'OSVersion' { $v = $v -replace '\s*\([\d.]+\)$', ''; break }
        'Office' { $v = $v -replace '\s*\(\d+\.\d+\.\d+\)$', ''; break }
        'Antivirus' { $v = $v -replace ', (aggiornato|firme non aggiornate)\)', ')'; break }
    }
    $v
}

function ConvertTo-HDChiaveProgramma {
    # Nome del programma senza numeri di versione: "7-Zip 24.08 (x64)" e "7-Zip 24.09 (x64)" sono lo stesso
    param([string]$Nome)
    ConvertTo-HDChiave (([string]$Nome) -replace '\d+([.,]\d+)+', '' -replace '(?i)\bupdate \d+\b', '')
}

function Format-HDElencoBreve {
    # Elenco di nomi in una cella: al massimo $Massimo, poi "e altri N"
    param([string[]]$Nomi, [int]$Massimo = 25)
    $n = @($Nomi | Where-Object { $_ })
    if ($n.Count -le $Massimo) { return ($n -join ', ') }
    '{0} e altri {1}' -f (($n | Select-Object -First $Massimo) -join ', '), ($n.Count - $Massimo)
}

function Format-HDCambiamento {
    # Frase dello storico per un campo cambiato: "RAM: 16 GB → 32 GB". Per i campi che sono elenchi solo
    # cosa è stato aggiunto, tolto o cambiato: "Monitor: aggiunto LG 27UL500; tolto DELL P2419H"
    param([string]$Campo, [string]$Prima, [string]$Dopo)
    $etichetta = $script:HDEtichetteStorico[$Campo]
    if (-not $etichetta) { $etichetta = $Campo }
    $separatore = $script:HDCampiElencoStorico[$Campo]
    if (-not $separatore) { return '{0}: {1} → {2}' -f $etichetta, $Prima, $Dopo }
    # separatore fuori dalle parentesi: "mario (disattivato, utente del PC)" è una voce sola
    $modello = '{0}\s*(?![^()]*\))' -f [regex]::Escape($separatore)
    $vecchi = @([regex]::Split([string]$Prima, $modello) | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $nuovi = @([regex]::Split([string]$Dopo, $modello) | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $chiave = { param($v) (($v -replace '(:\s.*|\s\(.*)$', '').Trim()).ToLowerInvariant() }
    $tolti = @($vecchi | Where-Object { $nuovi -notcontains $_ })
    $parti = New-Object System.Collections.Generic.List[string]
    foreach ($a in @($nuovi | Where-Object { $vecchi -notcontains $_ })) {
        $k = & $chiave $a
        $stesso = @($tolti | Where-Object { (& $chiave $_) -eq $k }) | Select-Object -First 1
        if ($stesso) {
            $parti.Add(('{0} → {1}' -f $stesso, $a))
            $tolti = @($tolti | Where-Object { $_ -ne $stesso })
        }
        else { $parti.Add('aggiunto ' + $a) }
    }
    foreach ($t in $tolti) { $parti.Add('tolto ' + $t) }
    if ($parti.Count -eq 0) { return '{0}: {1} → {2}' -f $etichetta, $Prima, $Dopo }
    '{0}: {1}' -f $etichetta, ($parti.ToArray() -join '; ')
}

function Get-HDCambiamenti {
    # Confronta la foto prima dei censimenti con le tabelle aggiornate: una frase per ogni valore cambiato
    # e per i programmi installati e disinstallati. Non conta il primo censimento di un PC (anche se la
    # riga c'era già dalla scansione) né una cella vuota che si riempie. Data = quella del censimento.
    param([Parameter(Mandatory)][hashtable]$Foto, [Parameter(Mandatory)]$Tabella, [datetime]$Adesso = (Get-Date))
    # data del censimento di ogni PC, dal foglio Dettagli aggiornato
    $date = @{}
    if ($Tabella.Dettagli) {
        foreach ($d in @($Tabella.Dettagli.Righe | Where-Object { $_ -and $_.Valori['Rilevato'] })) {
            $s = Get-HDNumeroSerie $d.Valori['SerialNumber']
            if ($s) { $date['S:' + $s] = [string]$d.Valori['Rilevato'] }
            $h = ConvertTo-HDNomeHost $d.Valori['Hostname']
            if ($h) { $date['H:' + $h] = [string]$d.Valori['Rilevato'] }
        }
    }
    $oggi = Format-HDData $Adesso
    $dataPc = {
        param($v)
        $s = Get-HDNumeroSerie $v['SerialNumber']
        if ($s -and $date.ContainsKey('S:' + $s)) { return $date['S:' + $s] }
        $h = ConvertTo-HDNomeHost $v['Hostname']
        if ($h -and $date.ContainsKey('H:' + $h)) { return $date['H:' + $h] }
        $oggi
    }
    $voce = {
        param($v, [string]$Testo)
        [pscustomobject]@{ Data = (& $dataPc $v); Hostname = [string]$v['Hostname']; SerialNumber = [string]$v['SerialNumber']; Cambiamento = $Testo }
    }
    $confronta = {
        param($Righe, [hashtable]$Prima, [string[]]$Campi, [switch]$SoloCensiti)
        foreach ($r in @($Righe | Where-Object { $_ })) {
            $vecchi = $Prima[[int]$r.Numero]
            if ($null -eq $vecchi) { continue }
            # riga arrivata dalla scansione e censita ora: è il primo censimento, non un cambiamento
            if ($SoloCensiti -and -not $vecchi['OS']) { continue }
            foreach ($k in $Campi) {
                $o = [string]$vecchi[$k]
                $n = [string]$r.Valori[$k]
                if (-not $o -or -not $n) { continue }
                if ((ConvertTo-HDValoreStorico $k $o) -ceq (ConvertTo-HDValoreStorico $k $n)) { continue }
                & $voce $r.Valori (Format-HDCambiamento -Campo $k -Prima $o -Dopo $n)
            }
        }
    }
    & $confronta $Tabella.Righe $Foto.Principale $script:HDCampiStorico -SoloCensiti
    if ($Tabella.Dettagli) { & $confronta $Tabella.Dettagli.Righe $Foto.Dettagli $script:HDCampiStoricoDettagli }
    if ($Tabella.DatiSoftware) {
        foreach ($r in @($Tabella.DatiSoftware.Righe | Where-Object { $_ })) {
            $vecchi = $Foto.Software[[int]$r.Numero]
            if ($null -eq $vecchi -or -not $vecchi['Programmi'] -or -not $r.Valori['Programmi']) { continue }
            if ([string]$vecchi['Programmi'] -ceq [string]$r.Valori['Programmi']) { continue }
            $prima = @{}
            foreach ($p in @(ConvertFrom-HDTestoProgrammi $vecchi['Programmi'])) { $prima[(ConvertTo-HDChiaveProgramma $p.Nome)] = $p.Nome }
            $dopo = @{}
            foreach ($p in @(ConvertFrom-HDTestoProgrammi $r.Valori['Programmi'])) { $dopo[(ConvertTo-HDChiaveProgramma $p.Nome)] = $p.Nome }
            $installati = @($dopo.Keys | Where-Object { -not $prima.ContainsKey($_) } | ForEach-Object { $dopo[$_] } | Sort-Object)
            $tolti = @($prima.Keys | Where-Object { -not $dopo.ContainsKey($_) } | ForEach-Object { $prima[$_] } | Sort-Object)
            if ($installati.Count -gt 0) { & $voce $r.Valori ('Programmi installati: ' + (Format-HDElencoBreve $installati)) }
            if ($tolti.Count -gt 0) { & $voce $r.Valori ('Programmi disinstallati: ' + (Format-HDElencoBreve $tolti)) }
        }
    }
}

function Add-HDStoricoATabella {
    # Cambiamenti -> righe nuove in fondo al foglio Storico
    param([Parameter(Mandatory)]$Tabella, [object[]]$Voci, [Parameter(Mandatory)][hashtable]$Modifiche)
    foreach ($v in @($Voci | Where-Object { $_ })) {
        $riga = New-HDRigaTabella $Tabella
        $diff = @{}
        foreach ($k in 'Data', 'Hostname', 'SerialNumber', 'Cambiamento') {
            $testo = [string]$v.$k
            if ($testo.Length -gt $script:HDMassimoCella) { $testo = $testo.Substring(0, $script:HDMassimoCella) }
            if ($testo -ne '') { $diff[$k] = $testo }
        }
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $Modifiche -Riga $riga -Diff $diff
    }
}

#endregion

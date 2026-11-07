# Censimento: riepilogo dei punti da sistemare nel parco PC (foglio "Riepilogo" dell'Excel, ricalcolato a
# ogni scrittura, e report HTML da consegnare) e storico dei cambiamenti tra un censimento e l'altro
# (foglio "Storico", una riga per cambiamento). Funzioni pure: lavorano sulle tabelle lette dall'Excel.

$script:HDFoglioRiepilogo = 'Riepilogo'
$script:HDFoglioStorico = 'Storico'

# Foglio Riepilogo: una riga per punto da sistemare (riscritto ogni volta, non va modificato a mano)
$script:HDColonneRiepilogo = @(
    New-HDColonnaFoglio 'Priorita' 'Priorità' 15
    New-HDColonnaFoglio 'Punto' 'Da controllare' 44
    New-HDColonnaFoglio 'NumeroPC' 'N. PC' 8 -Numerica
    New-HDColonnaFoglio 'PC' 'PC' 60
    New-HDColonnaFoglio 'CosaFare' 'Cosa fare' 70
)

# Foglio Storico: una riga per cambiamento, si aggiunge in fondo a ogni censimento
$script:HDColonneStorico = @(
    New-HDColonnaFoglio 'Data' 'Data' 17
    New-HDColonnaFoglio 'Hostname' 'Hostname' 18
    New-HDColonnaFoglio 'SerialNumber' 'Serial Number' 16
    New-HDColonnaFoglio 'Campo' 'Cosa è cambiato' 26
    New-HDColonnaFoglio 'Prima' 'Prima' 46
    New-HDColonnaFoglio 'Dopo' 'Dopo' 46
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

function Get-HDPuntiRiepilogo {
    # Punti da sistemare nel parco PC, dai più urgenti: Priorita (Urgente, Da pianificare, Da sapere), Punto,
    # CosaFare e PC (Nome, Utilizzatore, Dettaglio). Solo i punti che riguardano almeno un PC.
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
            Priorita = 'Urgente'; Punto = 'Dischi a rischio di guasto'
            Cosa     = 'Fare subito una copia dei dati e sostituire il disco.'
            Prova    = { param($p) ([string]$p.Dettagli['Smart'] + ' ' + [string]$p.Dettagli['DischiFisici']) -cmatch 'GUASTO|ATTENZIONE' }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Windows non riceve più aggiornamenti di sicurezza'
            Cosa     = 'Passare a Windows 11 i PC compatibili (l''aggiornamento è gratuito) e sostituire gli altri; nel frattempo si possono acquistare gli aggiornamenti estesi (ESU).'
            Prova    = {
                param($p)
                $vecchio = Get-HDWindowsSenzaAggiornamenti -SistemaOperativo $p.Valori['OS'] -Versione $p.Valori['OSVersion'] -Adesso $Adesso
                if (-not $vecchio -or $vecchio -like 'Windows 11*') { return $false }
                $w11 = [string]$p.Dettagli['Windows11']
                if ($w11 -match '^S') { return "$vecchio, compatibile con Windows 11" }
                if ($w11 -match '^No') { return "$vecchio, non compatibile con Windows 11" }
                $vecchio
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Versione di Windows 11 non più aggiornata'
            Cosa     = 'Installare l''ultima versione di Windows 11 da Windows Update (è gratuita).'
            Prova    = {
                param($p)
                $vecchio = Get-HDWindowsSenzaAggiornamenti -SistemaOperativo $p.Valori['OS'] -Versione $p.Valori['OSVersion'] -Adesso $Adesso
                if ($vecchio -like 'Windows 11*') { return $vecchio }
                $false
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Antivirus assente o spento'
            Cosa     = 'Installare o riattivare l''antivirus e controllare che si aggiorni.'
            Prova    = {
                param($p)
                $av = [string]$p.Dettagli['Antivirus']
                if (-not $av) { return $false }
                if ($av -notmatch '\(attivo') { return $av }
                $false
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Cynet non installato'
            Cosa     = 'Installare l''agente Cynet.'
            # solo dove Cynet è in uso (almeno un PC lo ha): per gli altri clienti non è un problema
            Prova    = { param($p) $usaCynet -and [string]$p.Valori['Cynet'] -eq 'No' }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Disco pieno'
            Cosa     = 'Liberare spazio subito: con il disco pieno Windows non si aggiorna e i programmi possono bloccarsi.'
            Prova    = {
                param($p)
                $pieni = @(Get-HDVolumiTesto $p.Valori['FreeDiskPct'] | Where-Object { $_.Pct -lt $s.DiscoCriticoPct -or $_.LiberoGB -lt $s.DiscoCriticoGB })
                if ($pieni.Count -eq 0) { return $false }
                (@($pieni | ForEach-Object { '{0} {1} GB liberi' -f $_.Unita, $_.LiberoGB }) -join ', ')
            }
        }
        @{
            Priorita = 'Urgente'; Punto = 'Windows non attivato'
            Cosa     = 'Attivare Windows con una licenza valida.'
            Prova    = { param($p) [string]$p.Dettagli['AttivazioneWindows'] -like 'Non attivato*' }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'Aggiornamenti di Windows in ritardo'
            Cosa     = 'Controllare Windows Update su questi PC: spazio su disco, errori in Impostazioni, servizio disattivato.'
            Prova    = {
                param($p)
                $ultimo = ConvertFrom-HDDataTesto $p.Dettagli['Aggiornamento']
                if (-not $ultimo) { return $false }
                if (((& $quando $p) - $ultimo).TotalDays -le $s.AggiornamentiGiorniAttenzione) { return $false }
                'ultimo il ' + $ultimo.ToString('dd/MM/yyyy', $script:HDInvariante)
            }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'Disco con poco spazio libero'
            Cosa     = 'Fare pulizia o spostare i dati: Windows Update ha bisogno di almeno 10-20 GB liberi.'
            Prova    = {
                param($p)
                $volumi = @(Get-HDVolumiTesto $p.Valori['FreeDiskPct'])
                if (@($volumi | Where-Object { $_.Pct -lt $s.DiscoCriticoPct -or $_.LiberoGB -lt $s.DiscoCriticoGB }).Count -gt 0) { return $false }
                $pochi = @($volumi | Where-Object { $_.Pct -lt $s.DiscoAttenzionePct -or $_.LiberoGB -lt $s.DiscoAttenzioneGB })
                if ($pochi.Count -eq 0) { return $false }
                (@($pochi | ForEach-Object { '{0} {1} GB liberi' -f $_.Unita, $_.LiberoGB }) -join ', ')
            }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'BitLocker non attivo'
            Cosa     = 'Attivare BitLocker, prima di tutto sui portatili, e conservare la chiave di ripristino in un posto sicuro: se il PC viene perso o rubato i dati restano protetti.'
            Prova    = {
                param($p)
                $b = [string]$p.Valori['BitLocker']
                if (-not $b) { return $false }
                $sistema = @($b -split ';' | Where-Object { $_.Trim() -match '^C:' }) | Select-Object -First 1
                if (-not $sistema) { $sistema = @($b -split ';')[0] }
                # "non attiva", "sospesa", "decrittografia in corso", "in attesa di attivazione": dati non protetti
                if ($sistema.Trim() -notmatch '^\w:\s*(non attiva|sospesa|decrittografia|in attesa)') { return $false }
                $batteria = [string]$p.Dettagli['Batteria']
                if ($batteria -and $batteria -ne 'Nessuna') { return 'portatile' }
                $true
            }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'L''utente lavora come amministratore del PC'
            Cosa     = 'Dare all''utente un account standard e usare un account separato per installare i programmi: limita molto i danni di virus e ransomware.'
            Prova    = { param($p) [string]$p.Dettagli['Amministratori'] -match 'utente del PC' }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'Office non più supportato'
            Cosa     = 'Passare a Microsoft 365 o a Office 2024: le versioni vecchie non ricevono più aggiornamenti di sicurezza.'
            Prova    = {
                param($p)
                $o = [string]$p.Dettagli['Office']
                if ($o -match '(?<!\d)(2007|2010|2013|2016|2019|2021)(?!\d)') { return 'Office ' + $Matches[1] }
                $false
            }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'Batteria del portatile usurata'
            Cosa     = 'Sostituire la batteria (sotto il 50% della capacità originale dura poco e può gonfiarsi).'
            Prova    = {
                param($p)
                $b = [string]$p.Dettagli['Batteria']
                if ($b -match '^(\d+)%.*(da sostituire|usurata)') { return $Matches[1] + '%' }
                $false
            }
        }
        @{
            Priorita = 'Da pianificare'; Punto = 'Poca memoria RAM'
            Cosa     = 'Aggiungere memoria se il PC lo permette, altrimenti metterlo in lista per la sostituzione.'
            Prova    = {
                param($p)
                $gb = Get-HDMemoriaGB $p.Valori['Memory']
                if ($gb -gt 0 -and $gb -lt ($s.RamMinimaGB - 0.5)) { return ('{0} GB' -f [math]::Round($gb, 1)) }
                $false
            }
        }
        @{
            Priorita = 'Da sapere'; Punto = 'Programmi di accesso remoto'
            Cosa     = 'Verificare che siano autorizzati e protetti (password, account aziendale) e disinstallare quelli che non servono.'
            Prova    = {
                param($p)
                $a = [string]$p.Dettagli['AccessoRemoto']
                if ($a -and $a -ne 'Nessuno') { return $a }
                $false
            }
        }
        @{
            Priorita = 'Da sapere'; Punto = 'BIOS fermo da anni'
            Cosa     = 'Controllare se il produttore ha aggiornamenti del BIOS (spesso correggono falle di sicurezza); se il PC è a fine vita, metterlo in lista per la sostituzione.'
            Prova    = {
                param($p)
                $d = ConvertFrom-HDDataTesto $p.Dettagli['Bios']
                if (-not $d) { return $false }
                if (((& $quando $p) - $d).TotalDays -lt 365.25 * $s.BiosAnni) { return $false }
                'BIOS del ' + $d.Year
            }
        }
        @{
            Priorita = 'Da sapere'; Punto = 'PC non censiti da tempo'
            Cosa     = 'Verificare se sono ancora in uso: se sono stati dismessi, scriverlo nelle Note.'
            Prova    = {
                param($p)
                $d = ConvertFrom-HDDataTesto $p.Dettagli['Rilevato']
                if (-not $d -or $d -gt $Adesso.AddMonths(-1 * [int]$s.CensimentoMesi)) { return $false }
                'ultimo censimento il ' + $d.ToString('dd/MM/yyyy', $script:HDInvariante)
            }
        }
    )
    foreach ($c in $controlli) {
        $colpiti = New-Object System.Collections.Generic.List[object]
        foreach ($p in @($PC | Where-Object { $_ })) {
            $esito = & $c.Prova $p
            if (-not $esito) { continue }
            $dettaglio = ''
            if ($esito -is [string]) { $dettaglio = $esito }
            $colpiti.Add([pscustomobject]@{ Nome = $p.Nome; Utilizzatore = $p.Utilizzatore; Dettaglio = $dettaglio })
        }
        if ($colpiti.Count -gt 0) {
            [pscustomobject]@{ Priorita = $c.Priorita; Punto = $c.Punto; CosaFare = $c.Cosa; PC = $colpiti.ToArray() }
        }
    }
}

function ConvertTo-HDRigheRiepilogo {
    # Punti del riepilogo -> righe del foglio Riepilogo (una per punto, poi la riga con data e numero di PC)
    param([object[]]$Punti, [int]$TotalePC, [datetime]$Adesso = (Get-Date))
    $elencoPunti = @($Punti | Where-Object { $_ })
    foreach ($p in $elencoPunti) {
        $elenco = (@($p.PC | ForEach-Object { if ($_.Dettaglio) { '{0} ({1})' -f $_.Nome, $_.Dettaglio } else { $_.Nome } }) -join ', ')
        if ($elenco.Length -gt $script:HDMassimoCella) { $elenco = $elenco.Substring(0, $script:HDMassimoCella - 20) + '... (elenco accorciato)' }
        , @($p.Priorita, $p.Punto, [string]@($p.PC).Count, $elenco, $p.CosaFare)
    }
    if ($elencoPunti.Count -eq 0) { , @('', 'Nessun punto da sistemare', '', '', '') }
    $nota = 'Calcolato il {0} su {1}: si aggiorna da solo a ogni censimento.' -f (Format-HDData $Adesso), (Format-HDQuantita $TotalePC 'PC censito' 'PC censiti')
    , @('', $nota, '', '', '')
}

function Get-HDRigheRiepilogoTabella {
    # Righe del foglio Riepilogo dai dati (già aggiornati in memoria) del foglio principale e di Dettagli
    param([Parameter(Mandatory)]$Tabella, [hashtable]$Soglie, [datetime]$Adesso = (Get-Date))
    $dettagli = @()
    if ($Tabella.Dettagli) { $dettagli = @($Tabella.Dettagli.Righe | Where-Object { $_ }) }
    $pc = @(Get-HDPcCensiti -Righe @($Tabella.Righe | Where-Object { $_ }) -Dettagli $dettagli)
    ConvertTo-HDRigheRiepilogo -Punti @(Get-HDPuntiRiepilogo -PC $pc -Soglie $Soglie -Adesso $Adesso) -TotalePC $pc.Count -Adesso $Adesso
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

function Get-HDCambiamenti {
    # Confronta la foto prima dei censimenti con le tabelle aggiornate: una voce per ogni valore cambiato
    # (solo se c'era già un valore: riempire una cella vuota non è un cambiamento), una per i PC nuovi e
    # una per i programmi installati e disinstallati. Data = quella del censimento del PC.
    param([Parameter(Mandatory)][hashtable]$Foto, [Parameter(Mandatory)]$Tabella, [datetime]$Adesso = (Get-Date))
    $titoli = @{}
    foreach ($c in @($script:HDColonneCensimento) + @($script:HDColonneDettagli)) { if (-not $titoli.ContainsKey($c.Chiave)) { $titoli[$c.Chiave] = $c.Titolo } }
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
        param($v, [string]$Campo, [string]$Prima, [string]$Dopo)
        [pscustomobject]@{ Data = (& $dataPc $v); Hostname = [string]$v['Hostname']; SerialNumber = [string]$v['SerialNumber']; Campo = $Campo; Prima = $Prima; Dopo = $Dopo }
    }
    $confronta = {
        param($Righe, [hashtable]$Prima, [string[]]$Campi, [switch]$Nuovi)
        foreach ($r in @($Righe | Where-Object { $_ })) {
            $vecchi = $Prima[[int]$r.Numero]
            # primo censimento del PC: riga nuova, oppure riga arrivata dalla scansione (senza sistema operativo)
            if ($Nuovi -and $r.Valori['OS'] -and ($null -eq $vecchi -or -not $vecchi['OS'])) {
                $descrizione = @((Join-HDProduttoreModello $r.Valori['Manufacturer'] $r.Valori['Model']), [string]$r.Valori['OS']) | Where-Object { $_ }
                & $voce $r.Valori 'Primo censimento' '' ($descrizione -join ', ')
                continue
            }
            if ($null -eq $vecchi) { continue }
            foreach ($k in $Campi) {
                $o = [string]$vecchi[$k]
                $n = [string]$r.Valori[$k]
                if (-not $o -or -not $n) { continue }
                if ((ConvertTo-HDValoreStorico $k $o) -ceq (ConvertTo-HDValoreStorico $k $n)) { continue }
                & $voce $r.Valori $titoli[$k] $o $n
            }
        }
    }
    & $confronta $Tabella.Righe $Foto.Principale $script:HDCampiStorico -Nuovi
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
            if ($installati.Count -gt 0) { & $voce $r.Valori 'Programmi installati' '' (Format-HDElencoBreve $installati) }
            if ($tolti.Count -gt 0) { & $voce $r.Valori 'Programmi disinstallati' (Format-HDElencoBreve $tolti) '' }
        }
    }
}

function Add-HDStoricoATabella {
    # Cambiamenti -> righe nuove in fondo al foglio Storico
    param([Parameter(Mandatory)]$Tabella, [object[]]$Voci, [Parameter(Mandatory)][hashtable]$Modifiche)
    foreach ($v in @($Voci | Where-Object { $_ })) {
        $riga = New-HDRigaTabella $Tabella
        $diff = @{}
        foreach ($k in 'Data', 'Hostname', 'SerialNumber', 'Campo', 'Prima', 'Dopo') {
            $testo = [string]$v.$k
            if ($testo.Length -gt $script:HDMassimoCella) { $testo = $testo.Substring(0, $script:HDMassimoCella) }
            if ($testo -ne '') { $diff[$k] = $testo }
        }
        Add-HDRigaInTabella -Tabella $Tabella -Modifiche $Modifiche -Riga $riga -Diff $diff
    }
}

#endregion

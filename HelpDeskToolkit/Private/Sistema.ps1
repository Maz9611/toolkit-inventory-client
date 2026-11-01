# Funzioni di supporto sui dati di Windows, usate da verifica postazione e censimento.
# Funzioni pure: si testano senza Windows.

function ConvertFrom-HDDataQfe {
    # Win32_QuickFixEngineering.InstalledOn è un testo: di solito M/d/yyyy,
    # a volte un FILETIME esadecimale, a volte vuoto.
    param($Valore, [datetime]$Adesso = (Get-Date))
    if ($null -eq $Valore) { return $null }
    if ($Valore -is [datetime]) { return $Valore }
    $s = ([string]$Valore).Trim()
    if (-not $s) { return $null }
    $d = [datetime]::MinValue
    $stile = [System.Globalization.DateTimeStyles]::None
    $formatiUsa = [string[]]@('M/d/yyyy', 'M/d/yyyy h:mm:ss tt', 'M/d/yyyy H:mm:ss')
    if ([datetime]::TryParseExact($s, $formatiUsa, $script:HDInvariante, $stile, [ref]$d)) {
        if ($d -le $Adesso.AddDays(1)) { return $d }
    }
    $formatiIta = [string[]]@('d/M/yyyy', 'd/M/yyyy H:mm:ss', 'yyyyMMdd', 'yyyy-MM-dd')
    if ([datetime]::TryParseExact($s, $formatiIta, $script:HDInvariante, $stile, [ref]$d)) {
        if ($d -le $Adesso.AddDays(1)) { return $d }
    }
    if ($s -match '^[0-9A-Fa-f]{15,16}$') {
        try { return [datetime]::FromFileTime([Convert]::ToInt64($s, 16)) } catch { Write-Verbose "FILETIME non valido: $s" }
    }
    $null
}

function Get-HDSupportoOS {
    param([string]$Caption)
    $c = [string]$Caption
    if ($c -match '(?i)windows (xp|vista|7|8)\b') {
        return [pscustomobject]@{ Stato = 'CRITICO'; Messaggio = 'Sistema operativo fuori supporto da anni: niente aggiornamenti di sicurezza. Pianifica sostituzione o aggiornamento.' }
    }
    if ($c -match '(?i)windows 10') {
        if ($c -match '(?i)LTS[BC]') {
            return [pscustomobject]@{ Stato = 'INFO'; Messaggio = 'Edizione LTSC/LTSB: ha un ciclo di vita proprio, verifica la sua data di fine supporto.' }
        }
        return [pscustomobject]@{ Stato = 'ATTENZIONE'; Messaggio = 'Windows 10 non è più supportato dal 14/10/2025 (salvo licenze ESU): pianifica il passaggio a Windows 11.' }
    }
    if ($c -match '(?i)windows 11') { return [pscustomobject]@{ Stato = 'OK'; Messaggio = '' } }
    if ($c -match '(?i)server') {
        return [pscustomobject]@{ Stato = 'INFO'; Messaggio = 'Sistema server: verifica il ciclo di vita della versione installata.' }
    }
    [pscustomobject]@{ Stato = 'INFO'; Messaggio = '' }
}

function Get-HDNumeroSerie {
    # Scarta i valori segnaposto che alcuni produttori lasciano nel BIOS
    param([string]$Valore)
    $s = ([string]$Valore).Trim()
    if ($s -match '^(?i)(to be filled.*|default string|system serial number|chassis serial number|base ?board serial number|serial number|not specified|not applicable|not available|none|n/?a|invalid|unknown|empty|oem|o\.e\.m\.?|0+|1+|x+|123456789|1234567890)$') { return '' }
    $s
}

function ConvertFrom-HDStatoAntivirus {
    # productState di SecurityCenter2: 2° byte = motore attivo, 3° byte = firme aggiornate
    param([int64]$ProductState)
    $hex = '{0:X6}' -f $ProductState
    $hex = $hex.Substring($hex.Length - 6)
    [pscustomobject]@{
        Attivo     = ($hex.Substring(2, 2) -in '10', '11')
        Aggiornato = ($hex.Substring(4, 2) -eq '00')
    }
}

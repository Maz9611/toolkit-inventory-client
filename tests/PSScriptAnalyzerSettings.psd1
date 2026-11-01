# Regole di PSScriptAnalyzer per il toolkit.
# Le regole di compatibilità verificano che il codice giri su Windows PowerShell 5.1
# (quello preinstallato su Windows 10/11) e su PowerShell 7.
@{
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        # Toolkit interattivo: i messaggi colorati a video sono voluti
        'PSAvoidUsingWriteHost'
        # I controlli della verifica postazione hanno tutti la stessa firma (Dati, Soglie, Adesso)
        # anche quando non usano ogni parametro
        'PSReviewUnusedParameter'
    )
    Rules        = @{
        PSUseCompatibleSyntax   = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.0')
        }
        PSUseCompatibleCommands = @{
            Enable         = $true
            TargetProfiles = @('win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework')
        }
        PSUseCompatibleTypes    = @{
            Enable         = $true
            TargetProfiles = @('win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework')
        }
    }
}

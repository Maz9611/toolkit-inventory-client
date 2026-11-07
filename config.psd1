# IMPOSTAZIONI DEL TOOLKIT
# Non devi modificare niente: il toolkit funziona già così.
# Cambia un valore solo se vuoi un comportamento diverso, poi salva (Blocco note va bene).
@{
    # Dove salvare i report della verifica (vuoto = cartella Report del toolkit)
    CartellaReport = ''

    # All'avvio propone di riaprire il programma come amministratore (serve per SMART, TPM e BitLocker)
    ProponiAmministratore = $true

    Censimento     = @{
        # Cartella degli Excel di censimento (vuoto = cartella Censimento del toolkit, cioè sulla chiavetta)
        Cartella         = ''

        # Excel da usare come modello per i censimenti nuovi (vuoto = Censimento\Modello.xlsx se esiste)
        ModelloExcel     = ''

        # Valori scritti nelle colonne Device Type e Dominio
        DeviceTypeClient = 'PC CLIENT'
        DeviceTypeServer = 'SERVER'
        DominioAssente   = 'No'

        # Testo cercato in processi, servizi e programmi per la colonna Cynet
        RilevaCynet      = 'Cynet'

        # Account che NON sono l'utente del PC (per esempio quello con cui entri tu):
        # il censimento non li scrive mai come "Nome Utente"
        AccountTecnici   = @('admin', 'administrator', 'amministratore', 'tecnico')

        # Quante copie di backup dell'Excel tenere (cartella Censimento\backup)
        BackupDaTenere   = 30
    }

    # Soglie della verifica postazione, dello spazio disco e del riepilogo del censimento
    Soglie         = @{
        DiscoAttenzionePct            = 10   # ATTENZIONE sotto il 10% libero...
        DiscoAttenzioneGB             = 20   # ...o sotto i 20 GB liberi
        DiscoCriticoPct               = 5
        DiscoCriticoGB                = 10
        UptimeGiorni                  = 14   # PC acceso da più di N giorni senza riavvio
        RamUsoPct                     = 90
        AggiornamentiGiorniAttenzione = 45
        AggiornamentiGiorniCritico    = 90
        FirmeAntivirusGiorni          = 7
        ScartoOrarioMinuti            = 5
        RamMinimaGB                   = 8    # riepilogo: "Poca memoria RAM" sotto questa quantità
        BiosAnni                      = 6    # riepilogo: "BIOS fermo da anni" oltre questa età
        CensimentoMesi                = 6    # riepilogo: "PC non censiti da tempo" oltre questi mesi
    }
}

# Toolkit Help Desk - regole del progetto

Modulo PowerShell per help desk (testi per l'utente in italiano), lanciato da chiavetta sui PC dei
clienti. Funzione principale: censimento PC su un Excel (.xlsx) con colonne fisse.

## Regole che non si discutono

- **Sola lettura sui PC**: niente pulizie, servizi, registro, file fuori dalle cartelle del toolkit
  (Censimento, Report). Le scritture passano da `SupportsShouldProcess` (-WhatIf).
- **Compatibile con Windows PowerShell 5.1** oltre che con PowerShell 7. Vietati: `??`, `?.`, ternario,
  `&&`/`||`, `ForEach-Object -Parallel`, `Join-String`, `Get-Uptime`, `ConvertFrom-Json -AsHashtable`,
  `-replace` con scriptblock, `Join-Path` con più di due parti, metodi .NET Core (`[Math]::Clamp`,
  `String.Contains(x, StringComparison)`, `[Text.Encoding]::Latin1`...). Niente `Invoke-WebRequest`.
  Attenzione a `-f` dentro le chiamate di metodo: `$x.Metodo(('{0}' -f $a, $b))` (doppie parentesi) e
  alle graffe letterali nelle stringhe di formato (`{{` `}}`).
- **File .ps1/.psm1/.psd1 in UTF-8 con BOM e CRLF**; niente virgolette tipografiche nel codice; `.cmd`
  solo ASCII. Lo verifica `tests/Qualita.Tests.ps1`.
- `@($null).Count` è 1: filtra sempre con `Where-Object { $_ }` prima di contare.
- I nomi delle variabili non distinguono maiuscole: una variabile locale con lo stesso nome di un
  parametro tipizzato (`$rete` e `[string]$Rete`) ne eredita il tipo. Usa nomi diversi.
- Le funzioni che restituiscono elenchi non usano `, @(...)` se il chiamante fa `@(...)` (array annidati).
- La regola PSUseSingularNouns di PSScriptAnalyzer su Windows segnala più casi che su Linux:
  niente nomi di funzione che finiscono in "s" (es. `...Antivirus`). In Windows PowerShell 5.1
  `Add-Type` non ha `-CompilerOptions`. La verifica vera è la build su GitHub Actions (Windows).

## Excel (Private/Xlsx.ps1)

- Nessuna libreria esterna: lo .xlsx è uno zip di XML (`System.IO.Compression`).
- `Update-HDXlsx` modifica solo le celle indicate: scrive stringhe "inline" (mai sharedStrings),
  numeri solo per le colonne numeriche, non tocca le celle con formule, mantiene lo stile esistente,
  le righe nuove copiano lo stile della riga sopra o della colonna; allarga tabelle, filtro e
  dimensione. Ordine di righe e celle sempre crescente.
- I comandi pubblici scrivono con `Save-HDCensimento`: copia temporanea, modifica, rilettura di
  controllo, backup in `Censimento\backup`, sostituzione.
- Validazione: i file prodotti devono passare l'Open XML SDK (`OpenXmlValidator`) senza errori nuovi.

## Censimento

- Colonne in `$script:HDColonneCensimento` (Private/CensimentoDati.ps1): ordine, titoli, sinonimi,
  colonne manuali (Utilizzatore, Note, Locazione). Nessuna colonna è numerica: Free Disk Space % e
  Logical Disk Size sono testo per disco, es. `(C:) 120 GB 25% (D:) 353 GB 53%`.
- La raccolta sul PC (`$script:HDRaccoltaCensimento`) è autonoma (gira anche via Invoke-Command): le
  letture lente (`$lente`: licenza, programmi, dischi/SMART) girano in parallelo in un RunspacePool,
  ognuna con testo autonomo, e i tempi finiscono in `Tempi` (record: `TempiRaccolta`); la
  conversione `ConvertTo-HDRecordCensimento` è pura e testata. Nei test si sostituisce
  `Invoke-HDRaccoltaInventario` con `Mock -ModuleName HelpDeskToolkit`.
- Abbinamento (`Find-HDRigaCensimento`): MAC (scartato se seriale o hostname lo contraddicono:
  docking station), poi seriale, poi hostname, poi IP solo su righe senza nome e MAC.
- Mai sovrascrivere con valori vuoti; colonne manuali solo da input esplicito; IP dalla scansione
  sempre aggiornato, il resto della scansione solo nelle celle vuote.
- Ogni censimento è salvato in JSON in `rilevazioni\<file>\da-inserire` prima di scrivere l'Excel;
  `Invoke-HDSincronizzazione` lo applica e lo archivia.
- Il formato dell'export di Advanced IP Scanner non è documentato: l'import riconosce le colonne
  per nome (`$script:HDCampiScansione`). Se arriva un export reale diverso, aggiungere i sinonimi
  e un test in `tests/Scansione.Tests.ps1`.

## Fogli Riepilogo e Storico (Private/CensimentoRiepilogo.ps1, Public/Riepilogo.ps1)

- Riepilogo: una riga per punto da sistemare (`Get-HDPuntiRiepilogo`, pura: controlli come scriptblock
  su `Get-HDPcCensiti` = righe del foglio principale + Dettagli). Riscritto a ogni `Save-HDCensimento`
  dai dati già aggiornati in memoria (`Get-HDRigheRiepilogoTabella`); stato con `Get-HDStatoFoglioRiscritto`
  (come Software). Date di fine supporto di Windows in `$script:HDFineSupportoWindows11` (da aggiornare
  quando escono versioni nuove); soglie in `Soglie` di config (RamMinimaGB, BiosAnni, CensimentoMesi).
- Storico: una riga per cambiamento, aggiunta in fondo. `Invoke-HDSincronizzazione` fa
  `Get-HDFotoCensimento` prima di applicare i censimenti e `Get-HDCambiamenti` dopo; campi seguiti in
  `$script:HDCampiStorico*` (niente IP, spazio libero, build, firme antivirus: cambiano troppo spesso).
- `Export-RiepilogoCensimento`: pagina HTML per il cliente in Report (stesso CSS del report di verifica).
  Mai "la tua azienda" o simili nei testi per il cliente.

## Fogli Dettagli e Software (Private/CensimentoDettagli.ps1)

- Colonne in `$script:HDColonneDettagli` e `$script:HDColonneSoftware`; i dati grezzi in più li raccoglie
  `$script:HDRaccoltaCensimento`, `ConvertTo-HDDettagliPC` (pura) li trasforma in testi e il record
  li porta in `Dettagli` (hashtable) e `Software` (elenco, `$null` se non raccolto).
- Dettagli: una riga per PC (seriale, poi hostname). DatiSoftware (nascosto): una riga per PC con
  l'elenco dei programmi in una cella (`nome<TAB>versione<TAB>editore` per riga, max 32000 caratteri).
  Software: una riga per programma (N. PC, "Tutti" / "Tutti tranne ..." / elenco, versioni),
  riscritto per intero da `Set-HDRigheFoglio` a ogni censimento partendo da DatiSoftware.
- Colonne nuove negli Excel esistenti: `Aggiungibile` (es. BitLocker nel foglio principale) e
  quelle mancanti in Dettagli vengono aggiunte in fondo all'intestazione (`Add-HDColonneMancanti`).
- `Save-HDCensimento -Extra` scrive tutti i fogli in un'unica copia temporanea; `Add-HDFoglioXlsx`
  crea i fogli mancanti negli Excel esistenti (aggiunge anche gli stili). Un foglio con lo stesso
  nome ma altro contenuto (`Conflitto`) non si tocca.
- Nei cicli di lettura e scrittura dell'Excel niente chiamate di funzione per cella (fogli grandi).
- Batteria (root/wmi, capacità attuale / di progetto), amministratori locali (ADSI WinNT sul gruppo
  S-1-5-32-544, letto anche da utente normale) e ultimo aggiornamento (Win32_QuickFixEngineering) sono
  nel foglio Dettagli; amministratori e aggiornamenti tra le letture lente in parallelo.
- PowerShell 7.6 ha un difetto: `@($lista)` su una `List[object]` lancia "Argument types do not match".
  Usa `@($lista | Where-Object { $_ })` o `.ToArray()`.

## Scansione di rete e finestra

- Scansione: ping, ARP, secondo giro di ping più lento per chi non ha risposto, poi nomi da NetBIOS,
  DNS inverso e mDNS: domanda diretta alla porta 5353 di ogni indirizzo e, solo per la rete del PC,
  domande al gruppo 224.0.0.251 (servizi comuni in `$script:HDServiziMdns` e nomi inversi a gruppi di 30)
  da una porta qualsiasi: i dispositivi rispondono in unicast ("legacy unicast"), quindi niente ascolto
  sulla 5353 (avviso del firewall). I nomi vengono dai record A e dai PTR inversi (`Get-HDNomiDaMulticast`).
  I client UDP si creano con `New-HDClientUdp` (su Windows disattiva SIO_UDP_CONNRESET). Produttore dal MAC con
  `HelpDeskToolkit/Dati/produttori-mac.txt.gz` (prefisso di 6, 7 o 9 cifre<TAB>nome: MA-L, MA-M e MA-S da
  oui-data/IEEE; si cerca prima il blocco più piccolo; nell'exe è una risorsa passata al modulo come
  `$script:HDTestoProduttoriMac`).

- `Private/ScansioneRete.ps1`: calcoli sulle reti, lettura ARP e NetBIOS (funzioni pure, testate) e le
  chiamate di rete vere (`Invoke-HDPingParallelo`, `Get-HDViciniArp`, `Resolve-HDNetbiosRemoto`,
  `Resolve-HDNomeDnsInverso`), che nei test si sostituiscono con Mock.
- `Interfaccia.ps1` (Windows Forms) è solo un guscio sopra i comandi del modulo: niente logica lì.
  Misure in pixel sempre con `Get-Px` / `New-Margini` (scala DPI: l'exe chiama `SetProcessDPIAware`).
  `-Verifica` costruisce la finestra e si chiude (usato da `Prova-Toolkit.ps1`). Gli errori vanno in
  `%TEMP%\HelpDeskToolkit-errori.txt`.

## ToolkitHelpDesk.exe (cartella exe)

- `exe/Avvio.cs` (C# 5, solo ASCII: deve compilare anche con il csc di .NET Framework) ospita
  Windows PowerShell 5.1 nel processo: carica il modulo come modulo dinamico (`New-Module`) con
  radice = cartella dell'exe, poi esegue `Interfaccia.ps1` nello stesso runspace STA. L'ospite ha
  `Name = 'ToolkitHelpDesk'` e nessuna UI: niente `Read-Host` (`Test-HDInterattivo` restituisce
  `$false`), `exit` arriva come codice di uscita dell'exe.
- `exe/Crea-Exe.ps1` (lanciato da `Crea-Exe.cmd`) unisce i file del modulo e la finestra, li comprime
  come risorse e compila: CodeDom con Windows PowerShell, Roslyn con PowerShell 7 (`-Riferimenti` fuori
  da Windows). `-SoloPacchetto` restituisce i testi senza compilare (usato da `tests/Exe.Tests.ps1`).
- Nel modulo e nella finestra non usare `$PSScriptRoot` o file accanto agli script se non con un
  ripiego: dentro l'exe non esistono. Dopo ogni modifica va ricreato l'exe.

## Comandi

```powershell
Invoke-Pester ./tests     # tutti i test (Pester 5), include PSScriptAnalyzer e controlli su BOM/encoding
```

Dopo ogni modifica: test verdi, poi `Avvia-Prova.cmd` su una macchina di prova Windows (utente
normale e amministratore) e apertura dell'Excel di prova in Microsoft Excel.

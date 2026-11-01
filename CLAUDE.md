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
- La raccolta sul PC (`$script:HDRaccoltaCensimento`) è autonoma (gira anche via Invoke-Command); la
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

## Scansione di rete e finestra

- `Private/ScansioneRete.ps1`: calcoli sulle reti, lettura ARP e NetBIOS (funzioni pure, testate) e le
  chiamate di rete vere (`Invoke-HDPingParallelo`, `Get-HDViciniArp`, `Resolve-HDNetbiosRemoto`,
  `Resolve-HDNomeDnsInverso`), che nei test si sostituiscono con Mock.
- `Interfaccia.ps1` (Windows Forms) è solo un guscio sopra i comandi del modulo: niente logica lì.
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

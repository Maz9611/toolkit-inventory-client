# Toolkit Help Desk

Programma da chiavetta per l'assistenza: **censimento dei PC su Excel** (con la scansione della rete
integrata), verifica rapida della postazione e spazio disco.

> *English summary: a portable help desk toolkit for Windows (interface in Italian). It builds a PC
> inventory in an Excel workbook (hardware, OS, network, users), discovers devices on the local
> network (ping, ARP, NetBIOS, DNS) and runs quick workstation checks. It only reads from the PCs:
> it writes nothing but its own Excel files and reports, next to the program. Single portable
> `ToolkitHelpDesk.exe` built from the PowerShell sources in this repository. MIT license.*

**Scarica:** l'ultima versione è nelle [Releases](../../releases): `ToolkitHelpDesk.zip` contiene il
programma, `helpdesk-toolkit.zip` il pacchetto completo con la prova.

> **Sui PC non cambia niente**: legge soltanto. Scrive solo i suoi file (l'Excel del censimento e i
> report) nella sua cartella, cioè sulla chiavetta.

## Come si usa

**Sulla chiavetta basta un file: `ToolkitHelpDesk.exe`.** Doppio clic e si apre la finestra del
toolkit. Il censimento (`Censimento\`) e i report (`Report\`) vengono salvati accanto all'exe,
quindi sulla chiavetta. Non serve installare niente: usa il Windows PowerShell già presente in
Windows 10 e 11.

> La prima volta Windows può mostrare "Windows ha protetto il PC" (l'exe non è firmato):
> **Ulteriori informazioni → Esegui comunque**. Per evitarlo: tasto destro sull'exe →
> Proprietà → spunta **Sblocca** prima di copiarlo sulla chiavetta. Antivirus ed EDR possono
> bloccarlo finché non è in elenco: vedi "Antivirus ed EDR" più sotto.

In alternativa, dalla cartella estratta dello zip (mai da dentro lo zip), **`Avvia-Toolkit.cmd`** apre
la stessa finestra senza exe.

La finestra ha tre schede:

| Scheda | Cosa fa |
|---|---|
| **Censisci questo PC** | Legge i dati del PC, ti fa scrivere Utilizzatore, Locazione e Note e con **Salva nell'Excel** aggiorna la riga del PC, o la aggiunge se non c'era. |
| **Scansione della rete** | **Scansiona la rete** trova i dispositivi accesi (10-20 secondi); **Aggiungi all'Excel** li mette nel censimento. Puoi anche importare un file di Advanced IP Scanner. |
| **Verifica postazione e dischi** | Controlli rapidi prima di un intervento e spazio libero sui dischi. |

In alto scegli quale Excel di censimento usare (**Scegli...**), ne crei uno nuovo per un cliente
(**Nuovo...**) o lo apri (**Apri in Excel**).

Un giro tipico da un cliente:
1. Dal tuo portatile, con la chiavetta: scheda **Scansione della rete** → **Scansiona la rete** →
   **Aggiungi all'Excel**. Il primo salvataggio chiede il nome del cliente e crea
   `Censimento\Censimento_<Cliente>.xlsx` sulla chiavetta.
2. Su ogni PC: chiavetta → `Avvia-Toolkit.cmd` → scheda **Censisci questo PC** → compila e salva.
   I PC che non erano in rete (spenti, portatili fuori sede) li accendi e fai lo stesso: vengono
   aggiunti in fondo.
3. A fine giro apri l'Excel dal tuo PC.

Scorciatoie senza finestra (dalla cartella estratta): **`Censisci-PC.cmd`** censisce subito il PC (fa
le domande nella finestra nera); **`Importa-Scansione.cmd`** importa un file di Advanced IP Scanner
(trascinaci sopra il file).

## La scansione nelle reti grandi

- Nella casella **Rete** puoi scrivere più reti separate da virgole (`192.168.1.0/24, 192.168.2.0/24`)
  o intervalli (`10.0.0.1-10.0.0.200`), fino a **4096 indirizzi** per volta. Una /24 richiede 10-20
  secondi, 4096 indirizzi circa un minuto.
- La rete proposta è quella del PC, intera fino a /20. Se la rete è più grande (es. /16) propone
  solo la /24 del PC: le altre parti le aggiungi tu.
- **Stessa rete del PC:** si trovano anche i dispositivi che non rispondono al ping (tabella ARP),
  con il MAC. Chi non risponde subito (telefoni e dispositivi a risparmio energetico) viene
  ricontattato in un secondo giro più lento.
- **Nomi:** NetBIOS (PC Windows), DNS e mDNS (router, stampanti, telefoni, dispositivi smart).
- **Produttore:** dal MAC, con l'elenco ufficiale IEEE dei produttori (incluso nel toolkit). I MAC
  "casuali" di telefoni e tablet vengono indicati come tali.
- **Altre reti/VLAN dietro un router:** si trovano solo i dispositivi che rispondono al ping, e senza
  MAC (l'ARP non attraversa i router). Molti PC Windows con il firewall attivo non rispondono al ping:
  li aggiunge comunque il censimento quando passi sul PC. I nomi arrivano da NetBIOS o dal DNS
  (in dominio di solito c'è).
- È lo stesso limite di Advanced IP Scanner: per una rete grande conviene scansionare dal PC che
  sta nella VLAN dei client, oppure ripetere la scansione da ogni VLAN.

## Antivirus ed EDR

- Una scansione di rete (questa come quella di Advanced IP Scanner) può far scattare un avviso di
  Cynet o di altri EDR. Avvisa prima chi gestisce la sicurezza del cliente.
- `ToolkitHelpDesk.exe` esegue PowerShell al suo interno: alcuni antivirus/EDR possono bloccarlo o
  metterlo in quarantena, soprattutto finché non è firmato. Se gestisci tu la console dell'EDR,
  aggiungi un'eccezione per l'exe. Vedi anche "Code signing policy".

## Le colonne dell'Excel

| Colonna | Cosa contiene |
|---|---|
| Hostname | nome del PC (scansione, poi il PC stesso) |
| Indirizzo IP | sempre aggiornato dall'ultima scansione o dal PC |
| Tipo IP | `DHCP` o `Statico` |
| MAC Address | scheda di rete con cui il PC è collegato |
| Device Type | `PC CLIENT`, oppure `SERVER` per i Windows Server |
| Manufacturer, Model | produttore e modello (per i Lenovo il nome commerciale, es. "ThinkPad T14 Gen 2i") |
| Nome Utente | utente collegato al PC (correggibile durante il censimento) |
| **Utilizzatore**, **Note** | li scrivi tu |
| Dominio | nome del dominio; `No` se il PC è in un gruppo di lavoro; `Entra ID (...)` se è collegato a Microsoft 365 |
| Serial Number | numero di serie dal BIOS, dal telaio o dal prodotto; nei PC assemblati quello della scheda madre, con davanti `MB` |
| Operating System, Operating System Version | es. `Windows 11 Pro` e `24H2 (26100.2033)` |
| Memory | es. `16 GB DDR4 SODIMM (2 x 8 GB)`, con tipo e formato se Windows li conosce |
| Processor Type | es. `Intel Core i5-10500 @ 3.10GHz` |
| Free Disk Space % | ogni disco con spazio libero e percentuale: `(C:) 120 GB 25% (D:) 353 GB 53%` |
| Logical Disk Size | dimensione di ogni disco: `(C:) 476 GB (D:) 666 GB` |
| Cynet | `Sì` o `No` (in rosso il "No") |
| BitLocker | `C: attiva`, `C: attiva; D: non attiva`, `C: sospesa`... Negli Excel di prima la colonna si aggiunge da sola in fondo |
| **Locazione** | la scrivi tu |

### Gli altri due fogli: Dettagli e Software

Ogni censimento sul PC scrive anche:

**Dettagli**: una riga per PC, ritrovata per numero di serie o hostname.

| Colonna | Esempio |
|---|---|
| Pronto per Windows 11 | `Sì`, `Già Windows 11`, `No: TPM 1.2 (serve 2.0), processore non supportato`, `Da verificare: ...` (in rosso i "No") |
| TPM, Secure Boot, Avvio | `2.0`, `Attivo` / `Disattivato`, `UEFI` / `BIOS` |
| Dischi fisici | `Samsung SSD 980 500GB (SSD NVMe, 500 GB, OK)` (in rosso `ATTENZIONE` e `GUASTO`) |
| SMART dei dischi | `Samsung SSD 980 500GB: usura 3%, 38 °C, 2.140 ore, nessun errore`; in rosso errori e `GUASTO PREVISTO`. **Solo avviando come amministratore** (altrimenti lo dice, senza cancellare quello già letto) |
| BIOS | versione e data, che dà anche un'idea dell'età del PC |
| Monitor | modello e numero di serie: `DELL P2419H (S/N CFV9N93, 2020)` |
| Stampanti | stampanti installate (non quelle virtuali come PDF e XPS), con l'IP se di rete |
| Office | `Microsoft 365 Apps for business (16.0.17126)`, `Office Home and Business 2019`, `Non installato` |
| Attivazione Windows | `Attivato (OEM, chiave nel BIOS)`, `Attivato (Volume KMS)`, `Non attivato` |
| Antivirus | `Windows Defender (attivo, aggiornato)` |
| Accesso remoto | `TeamViewer; AnyDesk; Desktop remoto di Windows attivo`, oppure `Nessuno` |
| Programmi installati, Rilevato il | quanti programmi e quando è stato fatto il censimento |

**Software**: **una riga per programma**, non per PC, così resta corto anche con tanti PC.

| Programma | Editore | Versioni | N. PC | PC |
|---|---|---|---|---|
| Google Chrome | Google LLC | `131.0 (38 PC); 130.0: PC-07, PC-12` | 40 | `Tutti` |
| Adobe Acrobat Reader | Adobe | `24.003` | 37 | `Tutti tranne PC-03, PC-09, PC-21` |
| AutoCAD 2024 | Autodesk | `24.3` | 2 | `PC-UFFICIOTECNICO, PC-05` |

- "Versioni" mostra subito chi è rimasto indietro; con il filtro su "PC" trovi i programmi di un PC.
- Componenti di sistema, aggiornamenti e librerie (Visual C++, .NET) non sono elencati.
- Il foglio si **ricostruisce a ogni censimento**: non scriverci dentro (le tue note andrebbero perse).
  Un programma disinstallato sparisce dal riepilogo.
- L'elenco completo dei programmi di ogni PC sta nel foglio nascosto **DatiSoftware** (una riga per
  PC): serve al toolkit per ricostruire il riepilogo, non cancellarlo.

- Negli Excel creati con le versioni precedenti fogli e colonne nuove si aggiungono da soli al primo
  censimento. Se hai già un tuo foglio chiamato "Dettagli" o "Software" con altro contenuto, il
  toolkit non lo tocca e te lo segnala.
- **Avviato come amministratore** legge TPM, BitLocker e SMART in modo preciso. Da utente normale
  stima TPM e BitLocker (dove non si può dirlo scrive "Da verificare") e non legge lo SMART.
- La lettura dei dati del PC richiede pochi secondi: le parti lente (licenza di Windows, programmi
  installati, dischi) vengono lette in parallelo. `Avvia-Prova.cmd` mostra quanto ci mette ogni parte.

### Le regole, per fidarsi del risultato

- **Come ritrova il PC nell'Excel.** Prima cerca il MAC (di qualsiasi scheda del PC), poi il
  numero di serie, poi il nome. Un MAC uguale non basta se seriale o nome dicono il contrario:
  così una docking station passata da un portatile all'altro non fa sovrascrivere la riga sbagliata.
- **Cosa non tocca mai.** Utilizzatore, Note e Locazione si scrivono solo se li compili tu. Un
  valore vuoto non cancella mai un dato già presente. Le celle con formule restano come sono.
- **Il tuo Excel puoi personalizzarlo.** Formattazione, colonne in più, filtri, tabelle, altri
  fogli, un titolo sopra le intestazioni: il toolkit trova da solo intestazioni e colonne per nome.
- **Prima di ogni scrittura fa un backup** in `Censimento\backup` (gli ultimi 30). Ogni scansione è
  salvata anche come CSV in `Censimento\scansioni`.
- **Nessun censimento va perso.** Se l'Excel non si può scrivere (per esempio perché è aperto in
  Excel), i dati restano da parte ed entrano al salvataggio successivo.
- **Hai già un tuo modello Excel?** Mettilo in `Censimento\Modello.xlsx`: i censimenti nuovi
  partono da quello.

## config.psd1 (impostazioni)

È il file delle impostazioni. **Non devi modificarlo**: il toolkit funziona già così. Serve solo
se un giorno vuoi cambiare un comportamento, per esempio scrivere `WORKSTATION` invece di
`PC CLIENT` in Device Type, o aggiungere ad `AccountTecnici` l'account con cui entri nei PC dei
clienti, così il censimento non lo scrive come "Nome Utente".

## La prova (solo quando arriva una versione nuova)

Non serve per usare il toolkit: è il controllo da fare **una volta, sulla tua macchina di prova**, ogni
volta che arriva una versione nuova, prima di portarla dai clienti.

1. Estrai lo zip completo.
2. Doppio clic su **`Avvia-Prova.cmd`**. Prova tutto sul PC vero:
   - i dati del PC, anche quelli per i fogli Dettagli e Software;
   - la scansione della rete;
   - un censimento completo su un Excel di prova;
   - la costruzione della finestra;
   - l'avvio di `ToolkitHelpDesk.exe`;
   - la verifica e lo spazio disco.

   Scrive solo in `%TEMP%\HelpDeskToolkit-Prova` e salva un log.
3. Alla fine apri l'Excel di prova: **Excel deve aprirlo senza il messaggio "Abbiamo riscontrato un
   problema in alcuni contenuti"**.
4. Ripeti con tasto destro → **Esegui come amministratore**.
5. Poi prova `ToolkitHelpDesk.exe` sulla tua rete e su un paio di PC.

Se qualcosa non va: il log della prova e gli errori della finestra
(`%TEMP%\HelpDeskToolkit-errori.txt`) sono quello che serve per correggere.

## Ricreare l'exe (solo se modifichi il codice)

Per usare il toolkit non serve: ogni versione nelle [Releases](../../releases) ha l'exe già pronto,
creato da GitHub Actions a partire da questo repository.

L'exe contiene una copia degli script: se modifichi a mano un file `.ps1`, l'exe che hai resta
quello vecchio. In quel caso doppio clic su **`Crea-Exe.cmd`**: usa il compilatore C# che c'è in ogni
Windows (non scarica né installa niente) e crea `ToolkitHelpDesk.exe` nella cartella del toolkit.
Poi rifai la prova. Un exe creato così non è firmato.

## Da riga di comando

```powershell
Import-Module .\HelpDeskToolkit
Invoke-ScansioneRete -NelCensimento -Cliente 'Studio Rossi'   # scansione + Excel
Invoke-ScansioneRete -Rete '10.0.1.0/24, 10.0.2.0/24'         # più reti
Add-PCCensimento -Utilizzatore 'Mario Rossi' -Locazione 'Reception'
Get-InventarioPC                                              # dati del PC, senza scrivere
Import-ScansioneRete .\scansione.csv                          # file di Advanced IP Scanner
Sync-Censimento                                               # dati rimasti in sospeso
Test-Postazione
Get-SpazioDisco
```

Ogni comando ha la guida: `Get-Help Invoke-ScansioneRete -Full`. Quelli che scrivono accettano `-WhatIf`.

## Problemi comuni

- **"L'argomento ... per il parametro -File non esiste"**: hai aperto il file da dentro lo zip.
  Estrai lo zip e riprova dalla cartella estratta.
- **"L'esecuzione di script è disabilitata"**: avvia sempre dai file `.cmd`. Se la policy è imposta
  dall'IT del cliente con i Criteri di gruppo, gli script vanno firmati.
- **La finestra non si apre**: il toolkit apre al suo posto il menu testuale. L'errore è in
  `%TEMP%\HelpDeskToolkit-errori.txt`.
- **Serial Number vuoto**: alcuni PC (soprattutto assemblati) non hanno il numero di serie nel BIOS.
  Il censimento mostra cosa c'è scritto al suo posto.
- **La scansione trova pochi dispositivi**: alcuni firewall bloccano ping e NetBIOS. I PC della
  stessa rete si trovano comunque tramite la tabella ARP; quelli dietro un router solo se rispondono
  al ping (vedi "La scansione nelle reti grandi").
- **L'exe non parte o sparisce subito**: probabilmente l'ha bloccato l'antivirus. Controlla la
  quarantena e `%TEMP%\HelpDeskToolkit-errori.txt`.

## Cosa è stato verificato e cosa no

- **Test automatici:** ne passano 335 (Excel con più fogli, scansione anche su più reti, abbinamenti,
  formati delle colonne, fogli Dettagli e Software, verifica postazione, motore dell'exe). A ogni
  modifica girano anche su Windows PowerShell 5.1 con GitHub Actions, che crea l'exe e lo avvia. PSScriptAnalyzer, con le regole di compatibilità per Windows PowerShell 5.1,
  non segnala problemi.
- **Excel prodotti:** superano il validatore ufficiale Open XML di Microsoft.
- **Provati su Windows:** la finestra, la scansione su una rete domestica, il censimento (seriale
  della scheda madre compreso); l'exe si crea e si avvia su Windows (GitHub Actions).
- **Non ancora provati su un PC vero:**
  - la lettura dei dati per i fogli Dettagli e Software (provata con dati simulati);
  - `ToolkitHelpDesk.exe` usato con la finestra visibile;
  - la scansione su più reti e su una rete aziendale grande.

  È quello che controlla `Avvia-Prova.cmd`.

## Code signing policy

Le versioni pubblicate nelle Releases sono create da GitHub Actions direttamente da questo
repository (`.github/workflows/build.yml`): test, creazione dell'exe e firma.

Free code signing provided by [SignPath.io](https://about.signpath.io), certificate by
[SignPath Foundation](https://signpath.org) *(dopo l'approvazione del progetto da parte di SignPath
Foundation; fino ad allora l'exe non è firmato)*.

- Committers and reviewers: [Maz9611](https://github.com/Maz9611)
- Approvers: [Maz9611](https://github.com/Maz9611)

**Privacy policy:** this program will not transfer any information to other networked systems unless
specifically requested by the user or the person installing or operating it. *Il programma non invia
dati a nessun servizio esterno: la scansione della rete contatta solo gli indirizzi indicati da chi la
avvia (ping, NetBIOS, DNS) e tutti i dati restano nei file Excel e nei report accanto al programma.*

## Licenza

MIT, vedi [LICENSE](LICENSE).

L'elenco dei produttori di schede di rete (`HelpDeskToolkit/Dati/produttori-mac.txt.gz`) deriva dal
registro pubblico IEEE MA-L tramite il pacchetto [oui-data](https://github.com/silverwind/oui-data)
(licenza BSD-2-Clause, vedi `HelpDeskToolkit/Dati/produttori-mac-LICENSE.txt`).

## Struttura

```
ToolkitHelpDesk.exe      il programma: l'unico file che serve sulla chiavetta
Crea-Exe.cmd             ricrea ToolkitHelpDesk.exe dopo una modifica
Avvia-Toolkit.cmd        apre la finestra senza exe
Censisci-PC.cmd          censisce subito questo PC, senza finestra
Importa-Scansione.cmd    importa un file di Advanced IP Scanner
Avvia-Prova.cmd          prova completa per la macchina di test
Interfaccia.ps1          la finestra
Toolkit.ps1              menu testuale (riserva se la finestra non si apre)
config.psd1              impostazioni (facoltative)
Censimento\              Excel, backup, scansioni e rilevazioni (creata al primo uso)
HelpDeskToolkit\         il modulo PowerShell
exe\                     sorgente dell'exe (Avvio.cs, icona, Crea-Exe.ps1)
tests\                   test automatici
```

# Sicurezza

## Download ufficiale

L'unica fonte ufficiale del Toolkit Help Desk è la pagina
[Releases](https://github.com/Maz9611/toolkit-inventory-client/releases) di questo repository.
Ogni versione è creata da GitHub Actions direttamente dal codice qui pubblicato
(`.github/workflows/build.yml`) e riporta l'impronta SHA256 di `ToolkitHelpDesk.exe`, nel testo della
release e nel file `SHA256.txt`.

Copie arrivate da altri siti, da fork, da allegati o da chiavette di altri **non sono ufficiali** e
potrebbero essere state modificate. Il codice è libero (licenza MIT): chiunque può farne una copia,
ma solo questo repository pubblica le versioni ufficiali.

### Come verificare l'exe

In PowerShell, nella cartella dell'exe:

```powershell
Get-FileHash .\ToolkitHelpDesk.exe -Algorithm SHA256
```

Il valore deve essere identico a quello indicato nella release da cui dici di averlo scaricato. Se è
diverso, non usare il file.

## Cosa fa (e cosa non fa) il programma

- Sui PC **legge soltanto**: non cambia impostazioni, servizi, registro o file di sistema.
- Scrive solo i suoi file (Excel del censimento, report, backup) nella sua cartella.
- Non invia dati a servizi esterni e non si aggiorna da solo. La scansione della rete contatta solo
  gli indirizzi della rete indicata (ping, alcune porte TCP comuni, NetBIOS, DNS, mDNS).

## Segnalare un problema di sicurezza

Se trovi una vulnerabilità, non pubblicarne i dettagli in una issue aperta: usa
**Security → Report a vulnerability** in questo repository. Per tutto il resto va bene una normale
issue.

---

*English: the only official downloads are the GitHub Releases of this repository, built by GitHub
Actions from this source; each release lists the SHA256 hash of `ToolkitHelpDesk.exe`. Verify it with
`Get-FileHash` before use. Please report vulnerabilities privately via Security → Report a
vulnerability.*

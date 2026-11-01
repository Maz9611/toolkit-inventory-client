// Toolkit Help Desk - il programma ToolkitHelpDesk.exe
//
// Contiene il modulo PowerShell e la finestra (Interfaccia.ps1), compressi come risorse, e li
// esegue con il Windows PowerShell 5.1 gia' presente in Windows: nessun file .ps1 su disco,
// nessuna finestra nera, nessun criterio di esecuzione da cambiare.
// Il censimento e i report vengono salvati nella cartella dell'exe (la chiavetta).
//
// Si crea con Crea-Exe.cmd (exe\Crea-Exe.ps1). Codice C# 5: deve compilare anche con il
// compilatore di .NET Framework 4 presente in ogni Windows.

using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Text;
using System.Threading;
#if !SENZA_FINESTRA
using System.Windows.Forms;
#endif

[assembly: AssemblyTitle("Toolkit Help Desk")]
[assembly: AssemblyProduct("Toolkit Help Desk")]
[assembly: AssemblyDescription("Censimento PC su Excel, scansione di rete, verifica postazione")]
[assembly: AssemblyCompany("@@AUTORE@@")]
[assembly: AssemblyCopyright("@@AUTORE@@")]
[assembly: AssemblyVersion("@@VERSIONE@@")]
[assembly: AssemblyFileVersion("@@VERSIONE@@")]

namespace ToolkitHelpDesk
{
    // Ospite minimo per PowerShell: niente console (UI nulla), raccoglie il codice di "exit".
    public sealed class Ospite : PSHost
    {
        private readonly Guid id = Guid.NewGuid();
        private readonly CultureInfo cultura = Thread.CurrentThread.CurrentCulture;
        private readonly CultureInfo culturaInterfaccia = Thread.CurrentThread.CurrentUICulture;
        private int codiceUscita;
        private bool uscitaRichiesta;

        public int CodiceUscita { get { return codiceUscita; } }
        public bool UscitaRichiesta { get { return uscitaRichiesta; } }

        public override string Name { get { return "ToolkitHelpDesk"; } }
        public override Version Version { get { return Assembly.GetExecutingAssembly().GetName().Version; } }
        public override Guid InstanceId { get { return id; } }
        public override PSHostUserInterface UI { get { return null; } }
        public override CultureInfo CurrentCulture { get { return cultura; } }
        public override CultureInfo CurrentUICulture { get { return culturaInterfaccia; } }

        public override void SetShouldExit(int exitCode)
        {
            codiceUscita = exitCode;
            uscitaRichiesta = true;
        }

        public override void EnterNestedPrompt() { throw new NotSupportedException(); }
        public override void ExitNestedPrompt() { throw new NotSupportedException(); }
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
    }

    // Carica il modulo ed esegue la finestra. Senza Windows Forms: si prova anche fuori da Windows.
    public static class Esecutore
    {
        public const string CaricaModulo =
            "param($Testo, $Radice) " +
            "New-Module -Name HelpDeskToolkit -ScriptBlock ([scriptblock]::Create($Testo)) -ArgumentList $Radice | Import-Module -Global";

        public static int Esegui(string modulo, string interfaccia, string radice, string percorsoExe, bool verifica, bool impostaCriterio)
        {
            if (impostaCriterio)
            {
                // Come "powershell.exe -ExecutionPolicy Bypass": vale solo per questo programma
                // (criterio "Process"). Un criterio imposto con i Criteri di gruppo resta comunque.
                Environment.SetEnvironmentVariable("PSExecutionPolicyPreference", "Bypass", EnvironmentVariableTarget.Process);
            }
            Ospite ospite = new Ospite();
            InitialSessionState stato = InitialSessionState.CreateDefault();
            using (Runspace spazio = RunspaceFactory.CreateRunspace(ospite, stato))
            {
                // La finestra deve girare nel thread principale (STA), come con "powershell.exe -STA"
                spazio.ApartmentState = ApartmentState.STA;
                spazio.ThreadOptions = PSThreadOptions.UseCurrentThread;
                spazio.Open();
                spazio.SessionStateProxy.SetVariable("HDPercorsoExe", percorsoExe);

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = spazio;
                    ps.AddScript(CaricaModulo).AddParameter("Testo", modulo).AddParameter("Radice", radice);
                    ps.Invoke();
                    if (ps.Streams.Error.Count > 0)
                    {
                        throw new InvalidOperationException("Modulo non caricato: " + ps.Streams.Error[0].ToString());
                    }
                }

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = spazio;
                    ps.AddScript(interfaccia);
                    if (verifica) { ps.AddParameter("Verifica", true); }
                    ps.Invoke();
                }
                return ospite.UscitaRichiesta ? ospite.CodiceUscita : 0;
            }
        }

        public static string LeggiRisorsa(string nome)
        {
            Stream flusso = Assembly.GetExecutingAssembly().GetManifestResourceStream(nome);
            if (flusso == null) { throw new InvalidOperationException("Risorsa mancante nell'exe: " + nome); }
            using (GZipStream gz = new GZipStream(flusso, CompressionMode.Decompress))
            using (StreamReader lettore = new StreamReader(gz, Encoding.UTF8))
            {
                return lettore.ReadToEnd();
            }
        }

        public static void ScriviErrore(string messaggio)
        {
            try
            {
                string file = Path.Combine(Path.GetTempPath(), "HelpDeskToolkit-errori.txt");
                File.AppendAllText(file, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "  " + messaggio + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception)
            {
                // Il log non e' indispensabile
            }
        }
    }

#if !SENZA_FINESTRA
    internal static class Programma
    {
        private const string Titolo = "Toolkit Help Desk";

        [STAThread]
        private static int Main(string[] argomenti)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object mittente, ThreadExceptionEventArgs e)
            {
                Segnala("Errore imprevisto: " + e.Exception.Message);
            };

            bool verifica = false;
            foreach (string a in argomenti)
            {
                if (a.TrimStart('/', '-').Equals("verifica", StringComparison.OrdinalIgnoreCase)) { verifica = true; }
            }

            string percorsoExe = Assembly.GetExecutingAssembly().Location;
            string radice = Path.GetDirectoryName(percorsoExe);

            if (!verifica && DentroUnoZip(radice))
            {
                MessageBox.Show(
                    "Hai aperto il programma da dentro lo zip.\n\n" +
                    "Prima estrai ToolkitHelpDesk.exe (tasto destro sullo zip > Estrai tutto) e copialo sulla chiavetta: " +
                    "il censimento viene salvato nella stessa cartella del programma.",
                    Titolo, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 3;
            }

            try
            {
                string modulo = Esecutore.LeggiRisorsa("modulo.ps1.gz");
                string interfaccia = Esecutore.LeggiRisorsa("interfaccia.ps1.gz");
                return Esecutore.Esegui(modulo, interfaccia, radice, percorsoExe, verifica, true);
            }
            catch (Exception ex)
            {
                if (!verifica) { Segnala("Il programma non si e' avviato: " + ex.Message); }
                else { Esecutore.ScriviErrore("Verifica dell'exe: " + ex.ToString()); }
                return 1;
            }
        }

        private static bool DentroUnoZip(string cartella)
        {
            // Se si fa doppio clic su un exe dentro uno zip, Windows lo estrae da solo nella cartella temporanea
            string temp = Path.GetFullPath(Path.GetTempPath()).TrimEnd('\\');
            string qui = Path.GetFullPath(cartella).TrimEnd('\\');
            return qui.StartsWith(temp + "\\", StringComparison.OrdinalIgnoreCase);
        }

        private static void Segnala(string messaggio)
        {
            Esecutore.ScriviErrore(messaggio);
            MessageBox.Show(messaggio + "\n\nDettagli in %TEMP%\\HelpDeskToolkit-errori.txt", Titolo, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
#endif
}

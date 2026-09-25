// Launcher.cs — executável mínimo que a Microsoft Store exige no pacote.
// Só abre o PainelSala.ps1 (que está na mesma pasta) com o PowerShell do Windows, sem console,
// e aponta a pasta de dados para %LOCALAPPDATA%\PainelSala (a pasta do app da Store é somente leitura).
// Compilado pelo Empacotar-MSIX.ps1 com o csc.exe que já vem no Windows (.NET Framework 4.x).
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string script = Path.Combine(appDir, "PainelSala.ps1");
            if (!File.Exists(script))
            {
                MessageBox.Show("Arquivo PainelSala.ps1 não encontrado em:\n" + appDir, "Painel de Sala", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }
            string dataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PainelSala");
            Directory.CreateDirectory(dataDir);

            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"System32\WindowsPowerShell\v1.0\powershell.exe");
            string extra = "";
            foreach (string a in args) extra += " " + a;   // permite -Demo, -Janela

            var psi = new ProcessStartInfo();
            psi.FileName = ps;
            psi.Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + script + "\"" + extra;
            psi.WorkingDirectory = appDir;
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.EnvironmentVariables["PAINELSALA_DATA"] = dataDir;
            psi.EnvironmentVariables["PAINELSALA_PACKAGED"] = "1";
            using (var p = Process.Start(psi))
            {
                p.WaitForExit();
                return p.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Não foi possível abrir o Painel de Sala:\n" + ex.Message, "Painel de Sala", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}

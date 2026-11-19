// Etapa 7: GESTEM.exe instalado por GESTEM.ps1 2.0.0 (Install-GestemExe). Vuelta a PowerShell desde control.json «exe»: "no":
// GESTEM.exe abre «GESTEM.ps1 -VolverRemoto» (Undo-GestemExe: los accesos directos como estaban) y se cierra.
using System;
using System.Diagnostics;
using System.IO;

namespace Gestem
{
    public static class ExeSwitch
    {
        public static string Dir = App.AppDir;          // pruebas: carpeta de instalación de prueba
        public static Action<string, string> StartHook;  // pruebas: qué se iba a abrir (programa, argumentos)

        // instalado por GESTEM.ps1 (no la copia de desarrollo de cs\): GESTEM.ps1 al lado y exe_instalado.txt con los accesos de antes
        public static bool Installed { get { return File.Exists(Path.Combine(Dir, "GESTEM.ps1")) && File.Exists(Path.Combine(App.DataDir, "exe_instalado.txt")); } }

        public static bool WantsPs(object remote) { return remote != null && Ps.Str(Json.Str(remote, "exe")).Trim() == "no" && Installed; }

        public static void Start(string file, string args)
        {
            if (StartHook != null) { StartHook(file, args); return; }
            ProcessStartInfo p = new ProcessStartInfo(file, args); p.UseShellExecute = false; p.CreateNoWindow = true; p.WorkingDirectory = Dir;
            Process.Start(p);
        }

        public static void BackToPs()
        {
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"System32\WindowsPowerShell\v1.0\powershell.exe");
            App.Log("control.json «exe»: \"no\": se vuelve a GESTEM (PowerShell)");
            Start(ps, "-NoProfile -ExecutionPolicy Bypass -STA -File \"" + Path.Combine(Dir, "GESTEM.ps1") + "\" -Oculto -VolverRemoto");
        }
    }
}

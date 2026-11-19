// Etapa 7: GESTEM.exe instalado por GESTEM.ps1 2.0.0 (Install-GestemExe). Vuelta a PowerShell desde control.json «exe»: "no":
// GESTEM.exe abre «GESTEM.ps1 -VolverRemoto» (Undo-GestemExe: los accesos directos como estaban) y se cierra.
// Se actualiza solo con app/exe_version.txt, app/GESTEM.exe y app/GESTEM.exe.sha256 (tools\publish_exe.ps1); el anterior queda en GESTEM.old.exe.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

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

        // solo la copia instalada (GESTEM.ps1 al lado) y con los datos de verdad: nunca la de desarrollo de cs\
        public static bool CanUpdate { get { return App.Real && File.Exists(Path.Combine(Dir, "GESTEM.ps1")); } }

        public static string Sha256(string file)
        {
            using (System.Security.Cryptography.SHA256 h = System.Security.Cryptography.SHA256.Create())
            using (FileStream fs = File.OpenRead(file)) return BitConverter.ToString(h.ComputeHash(fs)).Replace("-", "");
        }

        // descarga, comprueba SHA-256 y --smoke; GESTEM.exe -> GESTEM.old.exe (Windows deja renombrar el que está abierto) y el nuevo en su sitio.
        // Si algo falla, GESTEM.exe queda como estaba (excepción con el motivo).
        public static void Update(string ver)
        {
            string exe = Path.Combine(Dir, "GESTEM.exe"), nw = Path.Combine(Dir, "GESTEM.new.exe"), old = Path.Combine(Dir, "GESTEM.old.exe");
            try
            {
                File.WriteAllBytes(nw, Updater.RepoBytes("app/GESTEM.exe", 120));
                string want = Ps.Str(Updater.RepoFile("app/GESTEM.exe.sha256", 20)).Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
                string got = Sha256(nw);
                if (want.Length == 0 || !string.Equals(got, want, StringComparison.OrdinalIgnoreCase)) throw new Exception("SHA-256 no coincide (" + got + ")");
                Smoke(nw);
                if (File.Exists(old)) File.Delete(old);
                File.Move(exe, old);
                try { File.Move(nw, exe); } catch { File.Move(old, exe); throw; }
            }
            finally { try { if (File.Exists(nw)) File.Delete(nw); } catch { } }
            App.Log("GESTEM.exe actualizado de " + App.Version + " a " + ver + " (el anterior queda en GESTEM.old.exe)");
        }

        static void Smoke(string file)
        {
            string log = Path.Combine(Dir, "smoke.log");    // --smoke lo escribe junto al programa
            if (File.Exists(log)) File.Delete(log);
            ProcessStartInfo si = new ProcessStartInfo(file, "--smoke"); si.UseShellExecute = false; si.WorkingDirectory = Dir;
            using (Process p = Process.Start(si))
            {
                if (!p.WaitForExit(60000)) { try { p.Kill(); } catch { } throw new Exception("--smoke no termina"); }
                bool ok = File.Exists(log) && File.ReadAllText(log).StartsWith("OK xaml");
                if (p.ExitCode != 0 || !ok) throw new Exception("--smoke falla (código " + p.ExitCode + ")");
            }
        }

        // abre el GESTEM.exe nuevo con los mismos argumentos; espera (--tras) a que este se cierre
        public static void Restart()
        {
            List<string> a = new List<string>();
            string[] all = Environment.GetCommandLineArgs();
            for (int i = 1; i < all.Length; i++) { if (all[i] == "--tras") { i++; continue; } a.Add(all[i].Contains(" ") ? "\"" + all[i] + "\"" : all[i]); }
            a.Add("--tras"); a.Add(Process.GetCurrentProcess().Id.ToString());
            Start(Path.Combine(Dir, "GESTEM.exe"), string.Join(" ", a.ToArray()));
        }

        public static void BackToPs()
        {
            string ps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"System32\WindowsPowerShell\v1.0\powershell.exe");
            App.Log("control.json «exe»: \"no\": se vuelve a GESTEM (PowerShell)");
            Start(ps, "-NoProfile -ExecutionPolicy Bypass -STA -File \"" + Path.Combine(Dir, "GESTEM.ps1") + "\" -Oculto -VolverRemoto");
        }
    }
}

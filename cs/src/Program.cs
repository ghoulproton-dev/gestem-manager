using System;
using System.IO;
using System.Windows;

namespace Gestem
{
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string log = Path.Combine(dir, "smoke.log");
            if (args.Length > 1 && args[0] == "--selftest")
            {
                try { return SelfTest.Run(args[1]); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(args[1], "cs_error.txt"), ex.ToString()); return 2; }
            }
            if (args.Length > 1 && args[0] == "--uitest")
            {
                try { return UiTest.Run(args[1], args.Length > 2 ? args[2] : ""); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(args[1], "ui_error.txt"), ex.ToString()); return 2; }
            }
            if (args.Length > 1 && args[0] == "--fbclean")    // borra lo que quedó en gestorias/elmahi/_tests de una prueba cortada (nada más)
            {
                try { return FbTest.Clean(args[1]); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(args[1], "cs_fbclean_error.txt"), ex.ToString()); return 2; }
            }
            if (args.Length > 2 && args[0] == "--fbtest")    // red de verdad, solo gestorias/elmahi/_tests/...; no entra en test.bat
            {
                try { return FbTest.Run(args[1], args[2]); }
                catch (Exception ex) { File.WriteAllText(Path.Combine(args[1], "cs_fb_error.txt"), ex.ToString()); return 2; }
            }
            try
            {
                if (args.Length > 0 && args[0] == "--smoke")
                {
                    Window w = MainWin.LoadXaml();
                    int named = 0;
                    foreach (string n in new string[] { "ClientsGrid", "Toast", "PanelForm", "NavClients", "SearchBox" })
                        if (w.FindName(n) != null) named++;
                    File.WriteAllText(log, "OK xaml=" + w.Title + " named=" + named + "/5 clr=" + Environment.Version);
                    return 0;
                }
                // INICIO: config, trámite de la última vez, ventana
                // hasta terminar el traslado: sin --real se trabaja con una COPIA (cs\data\prueba) y documentos de prueba (cs\data\docs)
                // instalado por GESTEM.ps1 2.0.0 (GESTEM.ps1 al lado): los datos de verdad sin --real; --prueba fuerza el modo prueba
                App.Real = Array.IndexOf(args, "--real") >= 0 || (Array.IndexOf(args, "--prueba") < 0 && File.Exists(Path.Combine(App.AppDir, "GESTEM.ps1")));
                // --tras <pid>: abierto por GESTEM.ps1 o por la actualización: primero se espera a que ese programa se cierre
                int tr = Array.IndexOf(args, "--tras"); int trPid;
                if (tr >= 0 && tr + 1 < args.Length && int.TryParse(args[tr + 1], out trPid)) { try { System.Diagnostics.Process.GetProcessById(trPid).WaitForExit(30000); } catch { } }
                // un solo GESTEM abierto (por usuario y modo): el segundo trae el primero al frente y se cierra sin tocar nada
                bool first;
                one = new System.Threading.Mutex(true, @"Local\GESTEM-MANAGER" + (App.Real ? "" : "-prueba"), out first);
                if (!first) { ShowOther(); return 0; }
                if (!App.Real) TestMode.Setup(); else Config.Load();
                App.EnsureDirs();
                AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e) { Exception x = e.ExceptionObject as Exception; if (x != null) App.LogError("programa", x); };
                Tx.Lang = Config.Cur.Language;
                G.T = Tramite.ById(Config.Cur.Tramite) ?? Tramite.All[0];
                Application app = new Application();
                app.Run(new MainWin().Win);    // el título dice MODO PRUEBA (ApplyBranding)
            }
            catch (Exception ex)
            {
                File.WriteAllText(log, "FAIL " + ex);
                App.LogError("inicio", ex);    // también al registro de GESTEM (antes solo quedaba en smoke.log)
                if (!(args.Length > 0 && args[0] == "--smoke")) MessageBox.Show(ex.GetType().Name + ": " + ex.Message + "\r\n\r\n" + App.LogPath, App.Name + " - error", MessageBoxButton.OK, MessageBoxImage.Error);
                return 1;
            }
            GC.KeepAlive(one);
            return 0;
        }

        static System.Threading.Mutex one;    // se suelta al cerrar el programa

        [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
        [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
        [System.Runtime.InteropServices.DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
        // shortcut: trae la primera ventana de GESTEM.exe que encuentra; si a la vez hay uno de prueba y otro --real puede traer el otro
        static void ShowOther()
        {
            System.Diagnostics.Process me = System.Diagnostics.Process.GetCurrentProcess();
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(me.ProcessName))
            {
                if (p.Id == me.Id || p.MainWindowHandle == IntPtr.Zero) continue;
                if (IsIconic(p.MainWindowHandle)) ShowWindow(p.MainWindowHandle, 9);    // SW_RESTORE
                SetForegroundWindow(p.MainWindowHandle);
                return;
            }
        }
    }
}

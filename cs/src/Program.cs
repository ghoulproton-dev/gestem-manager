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
                App.EnsureDirs();
                AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e) { Exception x = e.ExceptionObject as Exception; if (x != null) App.LogError("programa", x); };
                Config.Load();
                Tx.Lang = Config.Cur.Language;
                G.T = Tramite.ById(Config.Cur.Tramite) ?? Tramite.All[0];
                Application app = new Application();
                app.Run(new MainWin().Win);
            }
            catch (Exception ex)
            {
                File.WriteAllText(log, "FAIL " + ex);
                App.LogError("inicio", ex);    // también al registro de GESTEM (antes solo quedaba en smoke.log)
                if (!(args.Length > 0 && args[0] == "--smoke")) MessageBox.Show(ex.GetType().Name + ": " + ex.Message + "\r\n\r\n" + App.LogPath, App.Name + " - error", MessageBoxButton.OK, MessageBoxImage.Error);
                return 1;
            }
            return 0;
        }
    }
}

using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Markup;

namespace Gestem
{
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            string log = Path.Combine(dir, "smoke.log");
            try
            {
                Application app = new Application();
                string xaml;
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("main.xaml"))
                using (StreamReader r = new StreamReader(s))
                    xaml = r.ReadToEnd();
                Window win = (Window)XamlReader.Parse(xaml);
                int named = 0;
                foreach (string n in new string[] { "ClientsGrid", "Toast", "PanelForm", "NavClients", "SearchBox" })
                    if (win.FindName(n) != null) named++;
                File.WriteAllText(log, "OK xaml=" + win.Title + " named=" + named + "/5 clr=" + Environment.Version);
                if (args.Length > 0 && args[0] == "--smoke") return 0;
                app.Run(win);
            }
            catch (Exception ex)
            {
                File.WriteAllText(log, "FAIL " + ex);
                return 1;
            }
            return 0;
        }
    }
}

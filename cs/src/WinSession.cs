// Etapa 4c: sesión (quién usa GESTEM en este PC). Copiado de Show-Session y Start-Session de GESTEM.ps1 (y sus eventos).
// shortcut: Start-Session no deja aún «actividad GESTEM - PC.txt» en la carpeta externa (Export-ActivityLog va con la copia externa, etapa 5).
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Gestem
{
    public partial class MainWin
    {
        void WireSession()
        {
            U<Button>("UserBtn").Click += delegate { ShowSession(); };
            U<Button>("SessGo").Click += delegate { StartSession(U<TextBox>("SessNew").Text); };
            U<TextBox>("SessNew").KeyDown += delegate(object s, KeyEventArgs e) { if (e.Key == Key.Return) StartSession(U<TextBox>("SessNew").Text); };
            U<TextBox>("SessNew").TextChanged += delegate { U("SessNewHint").Visibility = U<TextBox>("SessNew").Text.Length > 0 ? Visibility.Collapsed : Visibility.Visible; };
        }

        static List<string> People()
        {
            List<string> l = new List<string>();
            foreach (string c in Config.Cur.Colleagues) if (Ps.Str(c).Trim().Length > 0) l.Add(c);
            return l;
        }

        public void ShowSession()
        {
            U<TextBlock>("SessTitle").Text = L("sessTitle"); U<TextBlock>("SessHint").Text = L("sessHint");
            U<ContentControl>("SessGo").Content = L("sessGo"); U<TextBlock>("SessNewHint").Text = L("sessNew");
            Panel names = U<Panel>("SessNames");
            names.Children.Clear();
            List<string> people = People();
            string last = Ps.Str(Config.Cur.LastUser);
            if (last.Length > 0)
            {
                List<string> l = new List<string>(); l.Add(last);
                foreach (string p in people) if (!string.Equals(p, last, StringComparison.CurrentCultureIgnoreCase)) l.Add(p);    // -ne de PowerShell
                people = l;
            }
            foreach (string p in people)
            {
                Button b = new Button();
                b.Style = (Style)Win.FindResource("Btn");
                b.Content = p; b.Tag = p;
                b.Height = 42; b.FontSize = 15; b.Padding = new Thickness(20, 0, 20, 0); b.Margin = new Thickness(0, 0, 8, 8);
                if (string.Equals(p, last, StringComparison.CurrentCultureIgnoreCase)) b.FontWeight = FontWeights.SemiBold;
                b.Click += delegate { StartSession((string)b.Tag); };
                names.Children.Add(b);
            }
            U<TextBox>("SessNew").Text = "";
            U("SessNewHint").Visibility = Visibility.Visible;
            U("PanelSession").Visibility = Visibility.Visible;
        }

        public void StartSession(string who)
        {
            who = Ps.Str(who).Trim().ToUpper();
            if (who.Length == 0) who = Ps.Str(Config.Cur.LastUser).Trim();
            if (who.Length == 0) return;
            App.User = who;
            Config.Cur.LastUser = who;
            List<string> people = People();
            if (!PsH.ContainsI(people, who)) people.Add(who);
            Config.Cur.Colleagues = people;
            Config.Save();
            string[] known; if (!G.Lists.TryGetValue("colleagues", out known) || known == null) known = new string[0];
            if (!PsH.ContainsI(new List<string>(known), who))
            {
                List<string> k = new List<string>(); foreach (string x in known) if (!string.IsNullOrEmpty(x)) k.Add(x); k.Add(who);
                G.Lists["colleagues"] = k.ToArray();
            }
            U("PanelSession").Visibility = Visibility.Collapsed;
            U<ContentControl>("UserBtn").Content = string.Format(L("userTip"), who);
            App.Log("sesión iniciada");
            string ext = Ps.Str(Config.Cur.ExtBackupDir).Trim();
            if (ext.Length > 0 && (Directory.Exists(ext) || File.Exists(ext))) ExtBackup.ExportActivityLog(ext);    // así se ve enseguida quién lo abrió
            if (EditRow == null) { BuildForm(); ResetForm(); }
            RefreshClientGrid();
        }
    }
}

// GESTEM.exe --uitest <carpeta> [escenario]: abre la ventana con una copia de los datos de prueba (<carpeta>\cs\x.xlsx +
// «GESTEM datos»), sin avisos que bloqueen, y la usa como una persona: cada trámite, cada cliente abierto en la ficha y
// cada campo cambiado. Cada error se cuenta y va al registro con la pila completa. Resultado: <carpeta>\ui_summary.txt.
// Escenarios: (nada) | fr | ar | canje0 (canje sin clientes) | recheck (vuelve a mirar las pestañas del Excel, con Excel)
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace Gestem
{
    public static class UiTest
    {
        public static int Run(string testDir, string scenario)
        {
            StringBuilder o = new StringBuilder();
            int errors = 0;
            App.Quiet = true; App.AddSafeRoot(testDir);
            App.DataDir = testDir; App.BackupDir = Path.Combine(testDir, "copias"); App.LogPath = Path.Combine(testDir, "ui_log.txt");
            App.CfgPath = Path.Combine(testDir, "config_ui.json"); FbState.Path = Path.Combine(testDir, "ui_firebase.json");
            Directory.CreateDirectory(App.BackupDir);
            Config.Cur = new Config(); Config.Cur.ExcelPath = Path.Combine(testDir, "cs", "x.xlsx");
            Config.Cur.UpdateKey = "clave-prueba"; Config.Cur.FbSync = false; Config.Cur.DocsDir = Path.Combine(testDir, "docs"); Config.Cur.Collaborators.Add("ABOGADO X");
            Config.Cur.TabsChecked = new List<string>();
            if (scenario != "recheck") foreach (Tramite t in Tramite.All) Config.Cur.TabsChecked.Add(t.Id);   // sin abrir Excel
            if (scenario == "fr" || scenario == "ar") Config.Cur.Language = scenario;
            Tx.Lang = Config.Cur.Language;
            if (scenario == "canje0")
            {
                // canje con el archivo de datos pero sin clientes (PC nuevo, gestoría que empieza)
                string p = Path.Combine(testDir, "cs", "GESTEM datos", "canje.json");
                StoreData s = Store.Read(Tramite.ById("canje")); s.Clients.Clear(); Store.Write(Tramite.ById("canje"), s);
                o.Append("UI canje0 clientes=" + Store.Read(Tramite.ById("canje")).Clients.Count + " " + File.Exists(p) + "\r\n");
            }
            G.T = Tramite.ById("canje");
            Application app = new Application();
            MainWin w = new MainWin();
            w.Win.ContentRendered += delegate
            {
                Action<string, Action> Try = delegate(string what, Action a)
                {
                    try { a(); }
                    catch (Exception e) { errors++; App.LogError("uitest " + what, e); if (errors <= 20) o.Append("UI! " + what + ": " + e.GetType().Name + ": " + e.Message + "\r\n"); }
                };
                Try("cabeceras", delegate
                {
                    List<string> h = new List<string>();
                    foreach (DataGridColumn c in w.U<DataGrid>("ClientsGrid").Columns) h.Add(Ps.Str(c.Header));
                    o.Append("UI cabeceras=" + string.Join("|", h.ToArray()) + "\r\n");
                    o.Append("UI estado=" + w.U<TextBlock>("StatusText").Text + "\r\n");
                });
                foreach (Tramite t in Tramite.All)
                {
                    Try(t.Id + " elegir", delegate { w.U<ComboBox>("TramiteBox").SelectedIndex = Tramite.All.IndexOf(t); Pump(w); });
                    Try(t.Id + " lista", delegate
                    {
                        w.U<ToggleButton>("NavClients").IsChecked = true;
                        DataGrid g = w.U<DataGrid>("ClientsGrid");
                        List<ClientRow> items = new List<ClientRow>(); foreach (object it in g.ItemsSource ?? new object[0]) items.Add((ClientRow)it);
                        int rows = G.Cache != null && G.Cache.ContainsKey(t.Id) ? G.Cache[t.Id].Rows.Count : -1;
                        // el mismo cliente dos veces: en GESTEM y leído otra vez de la pestaña hecha a mano
                        HashSet<string> gest = new HashSet<string>(); int dup = 0;
                        foreach (ClientRow r in items) if (r.Row < Hand.RowBase) gest.Add(Hand.HandKey(r.Name));
                        foreach (ClientRow r in items) if (r.Row >= Hand.RowBase && gest.Contains(Hand.HandKey(r.Name))) dup++;
                        o.Append("UI " + t.Id + " grid=" + items.Count + " rows=" + rows + (items.Count == rows ? " OK" : " MAL") + " repetidos=" + dup + "\r\n");
                        // cada cliente abierto en la ficha (como doble clic) y cada campo cambiado (como al escribir)
                        foreach (ClientRow r in items)
                        {
                            g.SelectedItem = r;
                            Try(t.Id + " " + r.Id + " editar", delegate { w.EditSelected(); });
                            Try(t.Id + " " + r.Id + " escribir", delegate { TouchAll(w); Pump(w); });
                        }
                        Try(t.Id + " nuevo", delegate { w.ResetForm(); TouchAll(w); w.ResetForm(); });
                    });
                }
                Try("buscar", delegate
                {
                    w.U<ComboBox>("TramiteBox").SelectedIndex = Tramite.All.IndexOf(Tramite.ById("canje"));
                    w.U<ToggleButton>("NavClients").IsChecked = true;
                    w.U<TextBox>("SearchBox").Text = "moh";
                    int want = 0; foreach (ClientRow r in G.Rows) if ((r.Name + " " + r.Doc + " " + r.Id + " " + r.Loc).ToLower().Contains("moh")) want++;
                    o.Append("UI buscar moh=" + w.FilteredRows().Count + " esperado=" + want + "\r\n");
                    w.U<TextBox>("SearchBox").Text = "";
                    w.U<ToggleButton>("OnlyDebt").IsChecked = true;
                    int debt = 0; foreach (ClientRow r in G.Rows) if (r.Pending > 0) debt++;
                    o.Append("UI deuda=" + w.FilteredRows().Count + " esperado=" + debt + "\r\n");
                    w.U<ToggleButton>("OnlyDebt").IsChecked = false;
                });
                Try("resumen", delegate
                {
                    // Resumen en la ventana de verdad: la gráfica se dibuja (el lienzo tiene tamaño), cambiar alcance, periodo y modo
                    w.U<ToggleButton>("NavStats").IsChecked = true; Pump(w);
                    int drawn = w.U<Canvas>("ChartCanvas").Children.Count;
                    w.U<ComboBox>("StatsScope").SelectedIndex = Tramite.All.IndexOf(Tramite.ById("canje")) + 1; Pump(w);
                    w.U<ComboBox>("StatsPeriod").SelectedIndex = 3; Pump(w);
                    foreach (string m in new string[] { "ChartDay", "ChartWeek", "ChartMonth" }) { w.U<RadioButton>(m).IsChecked = true; Pump(w); }
                    w.U<Button>("StatsAllBtn").RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent)); Pump(w);
                    o.Append("UI resumen grafica=" + (drawn > 0 ? "OK" : "MAL") + " filas=" + w.U<Grid>("StatsGrid").RowDefinitions.Count + " alcance=" + (w.StatsScopeId ?? "todos") + "\r\n");
                    w.U<ToggleButton>("NavClients").IsChecked = true;
                });
                Try("ajustes", delegate
                {
                    w.U<ToggleButton>("NavSettings").IsChecked = true; Pump(w);
                    w.U<ToggleButton>("NavAbout").IsChecked = true; Pump(w);
                    o.Append("UI ajustes nat=" + w.U<ComboBox>("NatBox").Items.Count + " logo=" + (w.U<System.Windows.Controls.Image>("AboutImg").Source != null ? "OK" : "MAL") + " ver=" + w.U<TextBlock>("AboutVer").Text + "\r\n");
                    w.U<ToggleButton>("NavClients").IsChecked = true;
                });
                Try("admin", delegate
                {
                    w.ShowAdminLogin(); Pump(w);
                    w.U<FrameworkElement>("PanelAdminLogin").Visibility = Visibility.Collapsed;
                    App.IsAdmin = true; w.OpenAdminPanel(); Pump(w);
                    foreach (string t in new string[] { "AdmTabLog", "AdmTabTools", "AdmTabLocks" }) { w.U<ToggleButton>(t).IsChecked = true; Pump(w); }
                    o.Append("UI admin panel=" + w.U<FrameworkElement>("PanelAdmin").Visibility + " registro=" + ((System.Collections.ICollection)w.U<ItemsControl>("AdmLog").ItemsSource).Count + "\r\n");
                    w.CloseAdmin(); Pump(w);
                });
                Try("aviso req", delegate { w.UpdateReqBadge(); o.Append("UI aviso req=" + w.U<Border>("ReqBadge").Visibility + "\r\n"); });
                Try("config", delegate
                {
                    // los ajustes de 1.8–1.9 sobreviven a un guardado (cambiar de trámite guarda config)
                    Config.Load();
                    Config c = Config.Cur;
                    o.Append("UI config " + (c.UpdateKey == "clave-prueba" && !c.FbSync && c.DocsDir == Path.Combine(testDir, "docs") && c.Collaborators.Contains("ABOGADO X") && c.TradProfit == 3 ? "OK" : "MAL") + "\r\n");
                });
                Pump(w);
                // también los que recogió el aviso de error de la ventana (eventos que llegan más tarde)
                int win = 0;
                try { foreach (string line in File.ReadAllLines(App.LogPath)) if (line.Contains("UI ERROR (ventana)")) { win++; if (win <= 5) o.Append("UI! " + line.Substring(Math.Min(21, line.Length)) + "\r\n"); } } catch (IOException) { }
                o.Append("UI errores=" + (errors + win) + "\r\n");
                File.WriteAllText(Path.Combine(testDir, "ui_summary.txt"), o.ToString(), new UTF8Encoding(false));
                w.Win.Close();
            };
            app.Run(w.Win);
            return 0;
        }

        // deja que lleguen los avisos pendientes de la ventana (como entre dos clics)
        static void Pump(MainWin w) { w.Win.Dispatcher.Invoke(new Action(delegate { }), System.Windows.Threading.DispatcherPriority.ContextIdle); }

        // como si alguien escribiera en cada campo de la ficha
        static void TouchAll(MainWin w)
        {
            foreach (Control c in w.FormControls())
            {
                if (!c.IsEnabled) continue;
                TextBox tb = c as TextBox; ComboBox cb = c as ComboBox; DatePicker dp = c as DatePicker; ContentControl cc = c as ContentControl;
                if (tb != null) tb.Text = tb.Text + "1";
                else if (cb != null && cb.IsEditable) cb.Text = "PRUEBA";
                else if (cb != null) { if (cb.Items.Count > 0) cb.SelectedIndex = cb.Items.Count - 1; }
                else if (dp != null) dp.SelectedDate = DateTime.Today;
                else if (cc != null && cc.Content is Panel) { Panel p = (Panel)cc.Content; if (p.Children.Count > 0) ((ToggleButton)p.Children[p.Children.Count - 1]).IsChecked = true; }
            }
        }
    }
}

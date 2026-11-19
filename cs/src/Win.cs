// Ventana principal (etapa 3a): paneles, trámite, lista de clientes, búsqueda y filtros, línea de estado, aviso (toast),
// aviso de requerimientos y carga al abrir. Copiado de Show-Panel / Apply-Language / Refresh-ClientGrid / Update-StatusLine /
// Show-Toast / Hide-Toast / Update-ReqBadge / Use-TramiteData y del bloque INICIO de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace Gestem
{
    public class MainWin
    {
        public Window Win;
        readonly Dictionary<string, FrameworkElement> ui = new Dictionary<string, FrameworkElement>();
        DispatcherTimer searchTimer, toastTimer;
        Action toastAction;
        bool switching;

        public static Window LoadXaml()
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("main.xaml"))
            using (StreamReader r = new StreamReader(s))
                return (Window)XamlReader.Parse(r.ReadToEnd());
        }

        // $script:UI.<nombre>: se busca una vez y se guarda
        public T U<T>(string name) where T : class
        {
            FrameworkElement e;
            if (!ui.TryGetValue(name, out e)) { e = Win.FindName(name) as FrameworkElement; ui[name] = e; }
            return e as T;
        }
        FrameworkElement U(string name) { return U<FrameworkElement>(name); }
        static Brush Brush(string hex) { return (Brush)new BrushConverter().ConvertFromString(hex); }
        static string L(string k) { return Tx.L(k); }

        public MainWin()
        {
            Win = LoadXaml();
            G.OnUseData = UseTramiteData;
            G.OnStatus = SetStatus;
            G.OnError = ShowError;

            U<ToggleButton>("NavClients").Checked += delegate { ShowPanel("PanelClients"); RefreshClientGrid(); U("SearchBox").Focus(); };
            U<ToggleButton>("NavNew").Checked += delegate { ShowPanel("PanelForm"); };
            U<ToggleButton>("NavStats").Checked += delegate { ShowPanel("PanelStats"); };
            U<ToggleButton>("NavColab").Checked += delegate { ShowPanel("PanelColab"); };
            U<ToggleButton>("NavSettings").Checked += delegate { ShowPanel("PanelSettings"); };
            U<ToggleButton>("NavAbout").Checked += delegate { ShowPanel("PanelAbout"); };
            U("ReqBadge").MouseLeftButtonUp += delegate
            {
                U<ComboBox>("TramiteBox").SelectedIndex = Tramite.All.IndexOf(Tramite.ById("requerimiento"));
                U<ToggleButton>("NavClients").IsChecked = true;
            };
            U<Button>("HelpBtn").Click += delegate { U("PanelHelp").Visibility = Visibility.Visible; };
            U<Button>("HelpClose").Click += delegate { U("PanelHelp").Visibility = Visibility.Collapsed; };

            // la lista se filtra cuando se deja de escribir un momento (escribir no se traba)
            searchTimer = new DispatcherTimer(); searchTimer.Interval = TimeSpan.FromMilliseconds(180);
            searchTimer.Tick += delegate { searchTimer.Stop(); RefreshClientGrid(); };
            U<TextBox>("SearchBox").TextChanged += delegate
            {
                U("SearchHint").Visibility = U<TextBox>("SearchBox").Text.Length > 0 ? Visibility.Collapsed : Visibility.Visible;
                searchTimer.Stop(); searchTimer.Start();
            };
            U<ComboBox>("FilterBox").SelectionChanged += delegate { RefreshClientGrid(); };
            U<ButtonBase>("OnlyDebt").Click += delegate { RefreshClientGrid(); };
            U<ButtonBase>("OnlyMine").Click += delegate { RefreshClientGrid(); };
            U<ComboBox>("TramiteBox").SelectionChanged += delegate { TramiteChanged(); };
            U<Button>("ToastBtn").Click += delegate { Action a = toastAction; HideToast(); if (a != null) a(); };
            U<Button>("ToastClose").Click += delegate { HideToast(); };
            Win.KeyDown += OnKeyDown;
            Win.Dispatcher.UnhandledException += delegate(object s, DispatcherUnhandledExceptionEventArgs e)
            {
                App.Log("UI ERROR: " + e.Exception);
                MessageBox.Show(e.Exception.Message, App.Name + " - error", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            };
            ApplyLanguage();
            Win.ContentRendered += delegate { Startup(); };
        }

        // --- avisos ---
        public void SetStatus(string msg)
        {
            U<TextBlock>("StatusText").Text = msg;
            Win.Dispatcher.Invoke(new Action(delegate { }), DispatcherPriority.Background);
        }
        public void ShowError(string msg)
        {
            App.Log("aviso mostrado: " + PsH.Rep(msg, @"\s+", " "));    // lo que ve la usuaria queda en «Errores» del admin
            MessageBox.Show(Win, msg, App.Name, MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        // aviso pequeño abajo a la derecha (sin ventana que haya que cerrar)
        public void ShowToast(string text, string btnText, Action action, int seconds)
        {
            U<TextBlock>("ToastText").Text = text;
            toastAction = action;
            Button b = U<Button>("ToastBtn");
            if (!string.IsNullOrEmpty(btnText) && action != null) { b.Content = btnText; b.Visibility = Visibility.Visible; } else b.Visibility = Visibility.Collapsed;
            U("Toast").Visibility = Visibility.Visible;
            if (toastTimer == null) { toastTimer = new DispatcherTimer(); toastTimer.Tick += delegate { HideToast(); }; }
            toastTimer.Stop(); toastTimer.Interval = TimeSpan.FromSeconds(seconds); toastTimer.Start();
        }
        public void HideToast()
        {
            if (toastTimer != null) toastTimer.Stop();
            toastAction = null;
            U("Toast").Visibility = Visibility.Collapsed;
        }

        // --- paneles ---
        public void ShowPanel(string name)
        {
            foreach (string p in new string[] { "PanelForm", "PanelClients", "PanelStats", "PanelColab", "PanelSettings", "PanelAbout" })
                U(p).Visibility = p == name ? Visibility.Visible : Visibility.Collapsed;
            // en Resumen el selector de arriba incluye «Todos los trámites»; la barra de trámite solo en Nuevo cliente, Clientes y Resumen
            bool onStats = name == "PanelStats";
            U("TopBar").Visibility = (name == "PanelForm" || name == "PanelClients" || onStats) ? Visibility.Visible : Visibility.Collapsed;
            U("TramiteBox").Visibility = onStats ? Visibility.Collapsed : Visibility.Visible;
            U("StatsScope").Visibility = onStats ? Visibility.Visible : Visibility.Collapsed;
        }

        // shortcut: solo los textos de las partes ya trasladadas (etapa 3a); el resto de Apply-Language llega con cada panel
        public void ApplyLanguage()
        {
            Tx.Lang = Config.Cur.Language;
            Win.FlowDirection = Config.Cur.Language == "ar" ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            U<ContentControl>("NavNew").Content = L("navNew"); U<ContentControl>("NavClients").Content = L("navClients"); U<ContentControl>("NavStats").Content = L("navStats");
            U<ContentControl>("NavSettings").Content = L("navSettings"); U<ContentControl>("NavAbout").Content = L("navAbout"); U<ContentControl>("NavColab").Content = L("navColab");
            U<ContentControl>("HelpBtn").Content = L("btnHelp");
            U<TextBlock>("TramiteLabel").Text = L("tramite");
            U<TextBlock>("SearchHint").Text = L("search"); U<ContentControl>("OnlyDebt").Content = L("onlyDebt"); U<ContentControl>("OnlyMine").Content = L("onlyMine");
            U<ContentControl>("RefreshBtn").Content = L("btnRefresh"); U<ContentControl>("OpenExcelBtn").Content = L("btnOpenExcel"); U<ContentControl>("EditBtn").Content = L("btnEdit");
            U<ContentControl>("DocsBtn").Content = L("btnDocs"); U<ContentControl>("ReceiptBtn").Content = L("btnReceipt");
            U<ContentControl>("ExportBtn").Content = L("btnExport") + "  " + (char)0x25BE;
            U<ContentControl>("UserBtn").Content = App.User.Length > 0 ? string.Format(L("userTip"), App.User) : L("sessGo");
            DataGrid grid = U<DataGrid>("ClientsGrid");
            string[] heads = { "colId", "colName", "colDoc", "colTipo", "colStatus", "colPay", "colPend", "colDays", "colResp" };
            for (int i = 0; i < heads.Length; i++) grid.Columns[i].Header = L(heads[i]);
            grid.Columns[grid.Columns.Count - 1].Header = L("colDocs");
            FillFilterBox();
            // trámites
            ComboBox tb = U<ComboBox>("TramiteBox");
            switching = true;
            try
            {
                tb.Items.Clear();
                foreach (Tramite t in Tramite.All) tb.Items.Add(t.TitleText());
                tb.SelectedIndex = Math.Max(0, Tramite.All.IndexOf(G.T));
            }
            finally { switching = false; }
            U<TextBlock>("FileText").Text = string.IsNullOrEmpty(Config.Cur.ExcelPath) ? "" : Path.GetFileName(Config.Cur.ExcelPath);
        }
        void FillFilterBox()
        {
            ComboBox fb = U<ComboBox>("FilterBox");
            fb.Items.Clear();
            fb.Items.Add(L("filterAll"));
            string[] st; if (G.Lists.TryGetValue("excel:N", out st)) foreach (string s in st) fb.Items.Add(s);
            fb.SelectedIndex = 0;
        }

        // --- trámite ---
        void TramiteChanged()
        {
            ComboBox tb = U<ComboBox>("TramiteBox");
            int i = tb.SelectedIndex;
            if (i < 0 || Tramite.All[i] == G.T) return;
            // mientras se carga un trámite no se puede cambiar a otro (si no, la ficha y el trámite se mezclan)
            if (switching) { tb.SelectedIndex = Tramite.All.IndexOf(G.T); return; }
            switching = true; Win.Cursor = Cursors.Wait;
            try
            {
                G.T = Tramite.All[i]; Config.Cur.Tramite = G.T.Id; Config.Save();
                G.Rows = null; U<DataGrid>("ClientsGrid").ItemsSource = null;
                if (Data.CachedData()) FillFilterBox();
            }
            finally { switching = false; Win.Cursor = null; }
        }

        // Use-TramiteData (lo llama Data.UseData)
        void UseTramiteData(TramiteData d)
        {
            RefreshClientGrid();
            UpdateStatusLine();
        }

        // --- lista de clientes ---
        public List<ClientRow> FilteredRows()
        {
            List<ClientRow> list = new List<ClientRow>();
            if (G.Rows == null) return list;
            string q = U<TextBox>("SearchBox").Text.Trim().ToLower();
            ComboBox fb = U<ComboBox>("FilterBox");
            string st = Ps.Str(fb.SelectedItem);
            bool onlyDebt = U<ToggleButton>("OnlyDebt").IsChecked == true, onlyMine = U<ToggleButton>("OnlyMine").IsChecked == true;
            foreach (ClientRow r in G.Rows)
            {
                if (q.Length > 0 && !(r.Name + " " + r.Doc + " " + r.Id + " " + r.Loc).ToLower().Contains(q)) continue;
                if (fb.SelectedIndex > 0 && r.Status != st) continue;
                if (onlyDebt && r.Pending <= 0) continue;
                if (onlyMine && App.User.Length > 0 && !Ps.EqI(r.Resp, App.User)) continue;
                // shortcut: nº de documentos (clip de la columna Docs) llega con los documentos (etapa 3d)
                r.DocsN = 0; r.DocsVis = "Collapsed";
                list.Add(r);
            }
            return list;
        }
        public void RefreshClientGrid()
        {
            if (G.Rows == null) return;    // lista vacía (trámite sin clientes) también se pinta
            U<DataGrid>("ClientsGrid").ItemsSource = FilteredRows();
        }
        public void UpdateStatusLine()
        {
            if (G.Rows == null) return;
            int n = 0; double sum = 0;
            foreach (ClientRow r in G.Rows) if (r.Pending > 0) { n++; sum += r.Pending; }
            SetStatus(string.Format(L("ready"), G.Rows.Count, n, Ps.Money(sum)));
        }

        // requerimientos abiertos que vencen en ≤3 días hábiles o ya vencidos; rojo si alguno venció
        public void UpdateReqBadge()
        {
            Border b = U<Border>("ReqBadge");
            int due = 0, late = 0;
            if (G.Cache != null && G.Cache.ContainsKey("requerimiento")) Store.ReqAlerts(G.Cache["requerimiento"].Rows, DateTime.Today, out due, out late);
            if (due + late == 0) { b.Visibility = Visibility.Collapsed; return; }
            U<TextBlock>("ReqBadgeText").Text = late > 0 ? string.Format(L("reqBadgeLate"), late, due) : string.Format(L("reqBadge"), due);
            b.Background = Brush(late > 0 ? "#B42318" : "#B54708");
            b.Visibility = Visibility.Visible;
        }

        void OnKeyDown(object sender, KeyEventArgs e)
        {
            bool ctrl = Keyboard.Modifiers == ModifierKeys.Control;
            if (ctrl && e.Key == Key.F)
            {
                U<ToggleButton>("NavClients").IsChecked = true; U("SearchBox").Focus(); U<TextBox>("SearchBox").SelectAll();
                e.Handled = true; return;
            }
            if (e.Key == Key.Escape && U("Toast").Visibility == Visibility.Visible) HideToast();
        }

        // al abrir (Add_ContentRendered): buscar el Excel, pestañas hechas a mano, cargar todo
        // shortcut: sin clave de GESTEM, actualizaciones, control remoto, documentos ni sesión (etapas 4, 5 y 5b)
        void Startup()
        {
            string x = Config.Cur.ExcelPath;
            if (string.IsNullOrEmpty(x) || !File.Exists(x))
            {
                SetStatus(L("noFile"));
                U<ToggleButton>("NavSettings").IsChecked = true;
                return;
            }
            Win.Cursor = Cursors.Wait;
            try
            {
                try { Hand.SyncTabs(true); } catch (Exception e) { App.Log("pestañas propias: " + e.Message); }
                if (Data.LoadData(true, false))
                {
                    FillFilterBox();
                    try { UpdateReqBadge(); } catch (Exception e) { App.Log("aviso de requerimientos: " + e.Message); }
                }
            }
            finally { Win.Cursor = null; }
            U<ToggleButton>("NavClients").IsChecked = true;
        }

        // GESTEM.exe --uitest <carpeta>: abre la ventana con los datos de prueba (cs\x.xlsx), comprueba la lista y se cierra
        public static int UiTest(string testDir)
        {
            System.Text.StringBuilder o = new System.Text.StringBuilder();
            App.DataDir = testDir; App.BackupDir = Path.Combine(testDir, "copias"); App.LogPath = Path.Combine(testDir, "ui_log.txt");
            App.CfgPath = Path.Combine(testDir, "config_ui.json"); FbState.Path = Path.Combine(testDir, "ui_firebase.json");
            Directory.CreateDirectory(App.BackupDir);
            Config.Cur = new Config(); Config.Cur.ExcelPath = Path.Combine(testDir, "cs", "x.xlsx");
            Config.Cur.UpdateKey = "clave-prueba"; Config.Cur.FbSync = false; Config.Cur.DocsDir = "D:/docs"; Config.Cur.Collaborators.Add("ABOGADO X");
            Config.Cur.TabsChecked = new List<string>(); foreach (Tramite t in Tramite.All) Config.Cur.TabsChecked.Add(t.Id);   // sin abrir Excel
            G.T = Tramite.ById("canje");
            Application app = new Application();
            MainWin w = new MainWin();
            w.Win.ContentRendered += delegate
            {
                try
                {
                    foreach (Tramite t in Tramite.All)
                    {
                        w.U<ComboBox>("TramiteBox").SelectedIndex = Tramite.All.IndexOf(t);
                        DataGrid g = w.U<DataGrid>("ClientsGrid");
                        int n = 0; foreach (object it in g.ItemsSource ?? new object[0]) n++;
                        int rows = G.Cache[t.Id].Rows.Count;
                        o.Append("UI " + t.Id + " grid=" + n + " rows=" + rows + (n == rows ? " OK" : " MAL") + " status=" + w.U<TextBlock>("StatusText").Text + "\r\n");
                    }
                    w.U<ComboBox>("TramiteBox").SelectedIndex = Tramite.All.IndexOf(Tramite.ById("canje"));
                    w.U<TextBox>("SearchBox").Text = "moh";
                    int want = 0; foreach (ClientRow r in G.Rows) if ((r.Name + " " + r.Doc + " " + r.Id + " " + r.Loc).ToLower().Contains("moh")) want++;
                    o.Append("UI buscar moh=" + w.FilteredRows().Count + " esperado=" + want + "\r\n");
                    w.U<TextBox>("SearchBox").Text = "";
                    w.U<ToggleButton>("OnlyDebt").IsChecked = true;
                    int debt = 0; foreach (ClientRow r in G.Rows) if (r.Pending > 0) debt++;
                    o.Append("UI deuda=" + w.FilteredRows().Count + " esperado=" + debt + "\r\n");
                    w.UpdateReqBadge();
                    // los ajustes de 1.8–1.9 sobreviven a un guardado (cambiar de trámite guarda config)
                    Config.Load();
                    Config c = Config.Cur;
                    o.Append("UI config " + (c.UpdateKey == "clave-prueba" && !c.FbSync && c.DocsDir == "D:/docs" && c.Collaborators.Contains("ABOGADO X") && c.TradProfit == 3 ? "OK" : "MAL") + "\r\n");
                    o.Append("UI aviso req=" + w.U<Border>("ReqBadge").Visibility + "\r\n");
                }
                catch (Exception e) { o.Append("UI! " + e + "\r\n"); }
                File.WriteAllText(Path.Combine(testDir, "ui_summary.txt"), o.ToString(), new System.Text.UTF8Encoding(false));
                w.Win.Close();
            };
            app.Run(w.Win);
            return 0;
        }
    }
}

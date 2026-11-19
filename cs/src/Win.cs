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
    public partial class MainWin
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
            U<ToggleButton>("NavStats").Checked += delegate { ShowPanel("PanelStats"); FillStatsFilters(); UpdateStats(); };

            U<ToggleButton>("NavSettings").Checked += delegate { LoadSettingsPanel(); ShowPanel("PanelSettings"); };
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
            WireForm();
            WireMenu();
            WireDocs(); WireStats(); WireSettings(); WireSession(); WireAdmin(); WireColab(); WireExt();
            Win.Dispatcher.UnhandledException += delegate(object s, DispatcherUnhandledExceptionEventArgs e)
            {
                App.LogError("ventana", e.Exception);
                if (!App.Quiet) MessageBox.Show(e.Exception.GetType().Name + ": " + e.Exception.Message + "\r\n\r\n" + App.LogPath, App.Name + " - error", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            };
            ApplyLanguage();
            InitBranding();
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
            if (ErrorHook != null) ErrorHook(msg);
            if (!App.Quiet) MessageBox.Show(Win, msg, App.Name, MessageBoxButton.OK, MessageBoxImage.Warning);
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
            U<TextBlock>("SumPriceL").Text = L("sumPrice"); U<TextBlock>("SumPaidL").Text = L("sumPaid"); U<TextBlock>("SumPendL").Text = L("sumPend");
            SetCmHeaders();
            SetDocsTexts();
            SetStatsTexts();
            SetSettingsTexts();
            SetColabTexts();            // Apply-Language (navColab va con los otros botones)
            ApplyAdminLanguage();
            BuildForm();
            ResetForm();
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
                if (Data.CachedData()) ApplyLanguage(); else { BuildForm(); ResetForm(); }
                if (StatsVisible) { StatsScopeId = G.T.Id; FillStatsFilters(); UpdateStats(); }
            }
            finally { switching = false; Win.Cursor = null; }
        }

        // Use-TramiteData (lo llama Data.UseData)
        void UseTramiteData(TramiteData d)
        {
            Docs.UpdateMap();
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
                int nd = Docs.Count(r);
                r.DocsN = nd; r.DocsVis = nd > 0 ? "Visible" : "Collapsed";
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
            Key key = e.Key == Key.System ? e.SystemKey : e.Key;
            ModifierKeys want = ModifierKeys.Control | ModifierKeys.Shift | ModifierKeys.Alt;
            if (key == Key.A && (Keyboard.Modifiers & want) == want)    // administrador (oculto)
            {
                if (App.IsAdmin) OpenAdminPanel(); else ShowAdminLogin();
                e.Handled = true; return;
            }
            if (U("PanelSession").Visibility == Visibility.Visible) return;    // eligiendo quién eres: sin atajos
            bool ctrl = Keyboard.Modifiers == ModifierKeys.Control;
            if (ctrl && e.Key == Key.F)
            {
                U<ToggleButton>("NavClients").IsChecked = true; U("SearchBox").Focus(); U<TextBox>("SearchBox").SelectAll();
                e.Handled = true; return;
            }
            if (ctrl && e.Key == Key.S && U("PanelForm").Visibility == Visibility.Visible) { SaveFromForm(); e.Handled = true; return; }
            if (ctrl && e.Key == Key.N)
            {
                U<ToggleButton>("NavNew").IsChecked = true;
                if (EditRow == null) FocusCtl("name");
                e.Handled = true; return;
            }
            if (ctrl && e.Key == Key.Z && !(Keyboard.FocusedElement is System.Windows.Controls.Primitives.TextBoxBase))
            {
                if (!Busy) InvokeQuickUndo();
                e.Handled = true; return;
            }
            if (e.Key == Key.Escape)
            {
                if (U("Toast").Visibility == Visibility.Visible) HideToast();
                if (U("PanelDocs").Visibility == Visibility.Visible) { U("PanelDocs").Visibility = Visibility.Collapsed; e.Handled = true; return; }
                if (U("PanelForm").Visibility == Visibility.Visible && EditRow != null) { ResetForm(); U<ToggleButton>("NavClients").IsChecked = true; e.Handled = true; return; }
            }
        }

        // al abrir (Add_ContentRendered): buscar el Excel, pestañas hechas a mano, cargar todo
        // shortcut: sin clave de GESTEM, actualizaciones, control remoto ni carpeta de documentos en la nube (etapas 4d, 5 y 5b)
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
                try { Docs.RepairAll(); } catch (Exception e) { App.Log("documentos: " + e.Message); }    // carpetas que quedaron con el nº de otro cliente
                try { Hand.SyncTabs(true, true); } catch (Exception e) { App.Log("pestañas propias: " + e.Message); }
                if (Data.LoadData(true, false))
                {
                    ApplyLanguage();
                    try { UpdateReqBadge(); } catch (Exception e) { App.Log("aviso de requerimientos: " + e.Message); }
                }
            }
            finally { Win.Cursor = null; }
            try { StartupExtBackup(); } catch (Exception e) { App.Log("copia externa: " + e.Message); }    // copia al abrir
            U<ToggleButton>("NavClients").IsChecked = true;
            // la última persona que usó este PC entra directamente; «¿No eres tú?» abre la lista de nombres
            string last = Ps.Str(Config.Cur.LastUser).Trim();
            if (last.Length > 0)
            {
                StartSession(last);
                ShowToast(string.Format(L("userHello"), last), L("userSwitch"), ShowSession, 8);
            }
            else ShowSession();
        }

    }
}

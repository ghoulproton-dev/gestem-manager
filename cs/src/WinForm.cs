// Ficha del cliente (etapa 3b). Copiado de New-FieldBox / Add-ToGrid / Connect-CalendarToday / Build-Form / Get-FieldValue /
// Set-FieldValue / Get-FormValues / Reset-Form / Fill-FormForEdit / On-FieldChanged / Update-Preview / Find-DuplicateDoc /
// Check-Duplicate / Save-FromForm / Get-EditConflict / Format-WhoWhen / Edit-Selected / Test-FormHasWork / Get-List.
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;

namespace Gestem
{
    public partial class MainWin
    {
        Dictionary<string, Control> ctl = new Dictionary<string, Control>();      // $script:Ctl
        Dictionary<string, WrapPanel> pills = new Dictionary<string, WrapPanel>(); // $script:PillGroups
        public ClientRow EditRow;
        bool filling, priceTouched, formDirty;
        static readonly Dictionary<string, string> PayColors = new Dictionary<string, string> { { "Pagado", "#A9D08E" }, { "Parcial", "#FFE699" }, { "Pendiente", "#FF9B9B" } };

        // Get-List: la lista sin huecos ni repetidos
        public static string[] GetList(string name)
        {
            name = name ?? "";
            if (name.StartsWith("fixed:", StringComparison.OrdinalIgnoreCase)) { List<string> f = new List<string>(); foreach (string x in name.Substring(6).Split('|')) if (x.Length > 0) f.Add(x); return f.ToArray(); }
            if (name == "collaborators") return CollaboratorNames();
            List<string> o = new List<string>();
            string[] l; if (G.Lists.TryGetValue(name, out l) && l != null) foreach (string x in l) if (x != null && x.Trim().Length > 0 && !o.Contains(x)) o.Add(x);
            return o.ToArray();
        }
        // Get-CollaboratorNames: los de Ajustes y los que ya se han escrito en cualquier trámite
        public static string[] CollaboratorNames()
        {
            List<string> set = new List<string>();
            foreach (string c in Config.Cur.Collaborators) { string x = (c ?? "").Trim().ToUpper(); if (x.Length > 0 && !set.Contains(x)) set.Add(x); }
            if (G.Cache != null)
                foreach (TramiteData d in G.Cache.Values)
                    foreach (ClientRow r in d.Rows) { string x = Ps.Str(Store.Get(r.Data, "colab")).Trim().ToUpper(); if (x.Length > 0 && !set.Contains(x)) set.Add(x); }
            string[] arr = set.ToArray(); Array.Sort(arr, StringComparer.CurrentCultureIgnoreCase);
            return arr;
        }

        void WireForm()
        {
            U<Button>("SaveBtn").Click += delegate
            {
                App.Log("clic Guardar");
                try { SaveFromForm(); } catch (Exception e) { App.Log("Guardar ERROR: " + e); ShowError(e.Message); }
            };
            U<Button>("ClearBtn").Click += delegate { bool wasEdit = EditRow != null; ResetForm(); if (wasEdit) U<ToggleButton>("NavClients").IsChecked = true; };
            DataGrid grid = U<DataGrid>("ClientsGrid");
            grid.MouseDoubleClick += delegate { if (grid.SelectedItem != null) EditSelected(); };
            U<Button>("EditBtn").Click += delegate { EditSelected(); };
            // Enter en la lista: abre el cliente seleccionado
            grid.PreviewKeyDown += delegate(object s, KeyEventArgs e) { if (e.Key == Key.Return && grid.SelectedItem != null) { EditSelected(); e.Handled = true; } };
            Win.Closing += delegate(object s, System.ComponentModel.CancelEventArgs e)
            {
                if (!FormHasWork()) return;
                string nl = Environment.NewLine + Environment.NewLine;
                if (MessageBox.Show(Win, string.Format(L("unsavedAsk"), nl), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                { e.Cancel = true; U<ToggleButton>("NavNew").IsChecked = true; }
            };
        }

        // ¿hay algo de verdad a medias en la ficha? (un cliente abierto para editar con cambios, o un nombre escrito)
        public bool FormHasWork()
        {
            if (!formDirty) return false;
            if (EditRow != null) return true;
            Control c; return ctl.TryGetValue("name", out c) && ((TextBox)c).Text.Trim().Length > 0;
        }

        StackPanel FieldBox(Field f, Control c)
        {
            StackPanel sp = new StackPanel(); sp.Margin = new Thickness(0, 0, 16, 14);
            TextBlock lbl = new TextBlock(); lbl.Text = f.LabelText(); lbl.Style = (Style)Win.FindResource("FieldLabel");
            sp.Children.Add(lbl); sp.Children.Add(c);
            return sp;
        }
        void AddToGrid(UIElement el, int col, int span)
        {
            Grid g = U<Grid>("FormGrid");
            if (col == 0) { RowDefinition rd = new RowDefinition(); rd.Height = GridLength.Auto; g.RowDefinitions.Add(rd); }
            Grid.SetRow(el, Math.Max(0, g.RowDefinitions.Count - 1)); Grid.SetColumn(el, col); Grid.SetColumnSpan(el, span);
            g.Children.Add(el);
        }
        // botón «Hoy» del calendario desplegable
        void ConnectCalendarToday(DatePicker dp)
        {
            try
            {
                Popup pop = dp.Template.FindName("PART_Popup", dp) as Popup;
                if (pop == null || pop.Child == null) return;
                Control cal = pop.Child as Control; if (cal == null) return;
                cal.ApplyTemplate();
                Button b = cal.Template.FindName("TodayBtn", cal) as Button;
                if (b == null) return;
                b.Content = L("calToday"); b.Tag = dp;
                if (b.Uid != "w")
                {
                    b.Uid = "w";
                    b.Click += delegate(object s, RoutedEventArgs e) { DatePicker d = (DatePicker)((Button)s).Tag; d.SelectedDate = DateTime.Today; d.DisplayDate = DateTime.Today; d.IsDropDownOpen = false; };
                }
            }
            catch (Exception e) { App.Log("calendario: " + e.Message); }
        }

        public void BuildForm()
        {
            Tramite T = G.T;
            Grid g = U<Grid>("FormGrid");
            g.Children.Clear(); g.RowDefinitions.Clear();
            ctl = new Dictionary<string, Control>(); pills = new Dictionary<string, WrapPanel>();
            string section = null; int col = 0;
            foreach (Field f in T.Fields)
            {
                if (f.Section != section)
                {
                    section = f.Section;
                    TextBlock secText = new TextBlock(); secText.Text = L("sec" + section); secText.Style = (Style)Win.FindResource("SectionTitle");
                    Border mark = new Border(); mark.Width = 4; mark.Height = 16; mark.CornerRadius = new CornerRadius(2); mark.Margin = new Thickness(0, 0, 10, 0);
                    mark.Background = Brush("#D8202A"); mark.VerticalAlignment = VerticalAlignment.Center;
                    StackPanel title = new StackPanel(); title.Orientation = Orientation.Horizontal; title.Margin = new Thickness(0, 18, 0, 12);
                    title.Children.Add(mark); title.Children.Add(secText);
                    AddToGrid(title, 0, 2); col = 0;
                }
                int span = f.Span == 2 ? 2 : 1;
                Control c;
                if (f.Type == "Combo")
                {
                    ComboBox cb = new ComboBox(); cb.Items.Add("");
                    foreach (string it in GetList(f.List)) cb.Items.Add(it);
                    cb.SelectionChanged += delegate(object s, SelectionChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                    c = cb;
                }
                else if (f.Type == "ComboEdit")
                {
                    ComboBox cb = new ComboBox(); cb.IsEditable = true; cb.IsTextSearchEnabled = true;
                    foreach (string it in GetList(f.List)) cb.Items.Add(it);
                    cb.SelectionChanged += delegate(object s, SelectionChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                    cb.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler(delegate(object s, TextChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); }));
                    c = cb;
                }
                else if (f.Type == "Date")
                {
                    DatePicker dp = new DatePicker(); dp.SelectedDateFormat = DatePickerFormat.Short;
                    // calendario en el idioma del programa (todos con fecha dd/mm/aaaa y semana desde el lunes)
                    string cal = Config.Cur.Language == "en" ? "en-GB" : Config.Cur.Language == "fr" ? "fr-FR" : Config.Cur.Language == "ar" ? "ar-MA" : "es-ES";
                    dp.Language = XmlLanguage.GetLanguage(cal);
                    dp.SelectedDateChanged += delegate(object s, SelectionChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                    dp.CalendarOpened += delegate(object s, RoutedEventArgs e) { ConnectCalendarToday((DatePicker)s); };
                    c = dp;
                }
                else if (f.Type == "Pills")
                {
                    WrapPanel wp = new WrapPanel();
                    string[] list = GetList(f.List);
                    if (list.Length == 0 && f.Key == "paystate") list = new string[] { "Pagado", "Parcial", "Pendiente" };
                    foreach (string it in list)
                    {
                        RadioButton rb = new RadioButton(); rb.Style = (Style)Win.FindResource("Pill");
                        rb.GroupName = "g_" + f.Key; rb.Content = it; rb.Tag = f.Key;
                        if (f.Colored && T.StatusColors.ContainsKey(it)) rb.Background = Brush(T.StatusColors[it]);
                        else if (f.Key == "paystate" && PayColors.ContainsKey(it)) rb.Background = Brush(PayColors[it]);
                        else rb.Uid = "plain";
                        rb.Checked += delegate(object s, RoutedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                        wp.Children.Add(rb);
                    }
                    pills[f.Key] = wp;
                    // el WrapPanel no es un Control: se envuelve para guardarlo en ctl (Tag e IsEnabled como en PowerShell)
                    ContentControl cc = new ContentControl(); cc.Content = wp; cc.Focusable = false; c = cc;
                }
                else if (f.Type == "Multiline")
                {
                    TextBox tb = new TextBox(); tb.Height = 80; tb.AcceptsReturn = true; tb.TextWrapping = TextWrapping.Wrap;
                    tb.VerticalContentAlignment = VerticalAlignment.Top; tb.Padding = new Thickness(8, 6, 8, 6);
                    tb.TextChanged += delegate(object s, TextChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                    c = tb;
                }
                else
                {
                    TextBox tb = new TextBox();
                    if (f.Type == "Upper") tb.CharacterCasing = CharacterCasing.Upper;
                    tb.TextChanged += delegate(object s, TextChangedEventArgs e) { OnFieldChanged((FrameworkElement)s); };
                    if (f.Key == "doc") tb.LostFocus += delegate { CheckDuplicate(); };
                    c = tb;
                }
                c.Tag = f.Key;
                if (f.ReadOnly) c.IsEnabled = false;    // calculado por GESTEM (fecha límite)
                ctl[f.Key] = c;
                StackPanel box = FieldBox(f, c);
                if (span == 2) { AddToGrid(box, 0, 2); col = 0; }
                else { AddToGrid(box, col, 1); col = 1 - col; }
            }
        }

        public object GetFieldValue(Field f)
        {
            Control c = ctl[f.Key];
            switch (f.Type)
            {
                case "Combo": return Ps.Str(((ComboBox)c).SelectedItem).Trim();
                case "ComboEdit": { string s = (((ComboBox)c).Text ?? "").Trim(); return f.KeepCase ? s : s.ToUpper(); }
                case "Date": { DateTime? d = ((DatePicker)c).SelectedDate; return d.HasValue ? (object)d.Value : null; }
                case "Money": { double? m = Ps.ParseMoney(((TextBox)c).Text); return m.HasValue ? (object)m.Value : null; }
                case "Pills": foreach (UIElement e in pills[f.Key].Children) { RadioButton rb = (RadioButton)e; if (rb.IsChecked == true) return Ps.Str(rb.Content); } return "";
                default: return ((TextBox)c).Text.Trim();
            }
        }
        public void SetFieldValue(Field f, object v)
        {
            Control c = ctl[f.Key];
            switch (f.Type)
            {
                case "Combo": { ComboBox cb = (ComboBox)c; string s = Ps.Str(v).Trim(); if (s.Length > 0 && !cb.Items.Contains(s)) cb.Items.Add(s); cb.SelectedItem = s; break; }
                case "ComboEdit": ((ComboBox)c).Text = Ps.Str(v).Trim(); break;
                case "Date":
                    if (v is DateTime) ((DatePicker)c).SelectedDate = (DateTime)v;
                    else if (v is double) ((DatePicker)c).SelectedDate = DateTime.FromOADate((double)v);
                    else ((DatePicker)c).SelectedDate = null;
                    break;
                case "Money": ((TextBox)c).Text = v is double ? ((double)v).ToString("0.##", Ps.ES) : ""; break;
                case "Pills": foreach (UIElement e in pills[f.Key].Children) { RadioButton rb = (RadioButton)e; rb.IsChecked = Ps.Str(rb.Content) == Ps.Str(v); } break;
                default: ((TextBox)c).Text = Ps.Str(v).Trim(); break;
            }
        }
        public Dictionary<string, object> GetFormValues()
        {
            Dictionary<string, object> vals = new Dictionary<string, object>();
            foreach (Field f in G.T.Fields) vals[f.Key] = GetFieldValue(f);
            return vals;
        }
        string V(Dictionary<string, object> v, string k) { return Ps.Str(Store.Get(v, k)); }

        public void ResetForm()
        {
            filling = true;
            EditRow = null; priceTouched = false;
            foreach (Field f in G.T.Fields)
            {
                object v = null;
                if (f.Role == "nationality") v = Config.Cur.DefaultNationality;
                else if (f.Role == "user") v = App.User;
                else if (f.DefaultToday) v = DateTime.Today;
                else if (f.Role == "gprice" && G.DefaultPrice.HasValue) v = G.DefaultPrice.Value;
                else if (!string.IsNullOrEmpty(f.Default)) v = f.Type == "Money" ? (object)Ps.Num(f.Default) : f.Default;    // en PowerShell el de importe es un número ([double]1)
                SetFieldValue(f, v);
            }
            filling = false;
            U<TextBlock>("FormTitle").Text = L("newTitle");
            U<ContentControl>("SaveBtn").Content = L("btnSave");
            U<ContentControl>("ClearBtn").Content = L("btnClear");
            U("DupWarn").Visibility = Visibility.Collapsed;
            UpdatePreview();
            formDirty = false;
            ctl["name"].Focus();
        }

        public void FillFormForEdit(ClientRow r)
        {
            filling = true;
            EditRow = r; priceTouched = true;
            foreach (Field f in G.T.Fields) SetFieldValue(f, Store.Get(r.Data, f.Key));
            filling = false; formDirty = false;
            U<TextBlock>("FormTitle").Text = string.Format(L("editTitle"), r.Id);
            U<ContentControl>("SaveBtn").Content = L("btnSaveEdit");
            U<ContentControl>("ClearBtn").Content = L("btnCancel");
            U("DupWarn").Visibility = Visibility.Collapsed;
            UpdatePreview();
        }

        void OnFieldChanged(FrameworkElement sender)
        {
            if (filling) return;
            formDirty = true;
            string key = Ps.Str(sender.Tag);
            Tramite T = G.T;
            if (key == "price") priceTouched = true;
            if ((key == "xF" || key == "xG") && ctl.ContainsKey("flim"))
            {
                DateTime? lim = Store.ReqDeadline(GetFieldValue(T.FieldByKey("xG")), GetFieldValue(T.FieldByKey("xF")));
                filling = true; SetFieldValue(T.FieldByKey("flim"), lim.HasValue ? (object)lim.Value : null); filling = false;
            }
            if (key == "tipo" && !priceTouched)
            {
                string tipo = Ps.Str(GetFieldValue(T.FieldByKey("tipo")));
                filling = true;
                if (tipo.Length > 0 && G.Tariff.ContainsKey(tipo)) ((TextBox)ctl["price"]).Text = G.Tariff[tipo].ToString("0.##", Ps.ES);
                else if (tipo.Length == 0) ((TextBox)ctl["price"]).Text = "";
                filling = false;
            }
            // traducciones: precio al cliente = presupuesto de la traductora + beneficio fijo por documento
            if (T.Id == "traduccion" && (key == "gastos" || key == "ndocs") && !priceTouched)
            {
                double? g = Ps.ParseMoney(((TextBox)ctl["gastos"]).Text);
                double? nd = Ps.ParseMoney(((TextBox)ctl["ndocs"]).Text); if (!nd.HasValue || nd.Value <= 0) nd = 1;
                if (g.HasValue) { filling = true; ((TextBox)ctl["price"]).Text = (g.Value + Config.Cur.TradProfit * nd.Value).ToString("0.##", Ps.ES); filling = false; }
            }
            UpdatePreview();
        }

        public void UpdatePreview()
        {
            Tramite T = G.T;
            Dictionary<string, object> v = GetFormValues();
            if (EditRow != null) { U<TextBlock>("CardId").Text = EditRow.Id; U<TextBlock>("CardTag").Text = ""; }
            else { U<TextBlock>("CardId").Text = G.NextId.Length > 0 ? G.NextId : "—"; U<TextBlock>("CardTag").Text = L("cardNew"); }
            string name = V(v, "name");
            U<TextBlock>("CardName").Text = name.Length > 0 ? name.ToUpper() : L("cardEmpty");
            U<TextBlock>("CardName").Foreground = Brush(name.Length > 0 ? "#14233A" : "#9AA7B8");
            List<string> sub = new List<string>(); foreach (string k in new string[] { "doc", "nat" }) if (V(v, k).Length > 0) sub.Add(V(v, k));
            U<TextBlock>("CardSub").Text = string.Join("  ·  ", sub.ToArray());
            U<TextBlock>("CardTipo").Text = V(v, "tipo").Length > 0 ? V(v, "tipo") : V(v, "xF");
            string status = V(v, "status");
            if (status.Length > 0)
            {
                U<TextBlock>("CardStatus").Text = status;
                U<Border>("CardStatusB").Background = Brush(T.StatusColors.ContainsKey(status) ? T.StatusColors[status] : "#E8EDF3");
                U("CardStatusB").Visibility = Visibility.Visible;
            }
            else U("CardStatusB").Visibility = Visibility.Collapsed;
            // pago
            double? price = Ps.Num(Store.Get(v, "price"));
            string pay = V(v, "paystate");
            Control paidBox; if (ctl.TryGetValue("paid", out paidBox)) paidBox.IsEnabled = pay == "Parcial";
            double? paid = pay == "Pagado" ? price : pay == "Parcial" ? Ps.Num(Store.Get(v, "paid")) : 0.0;
            double? pend = null;
            if (price.HasValue) pend = Math.Max(0, price.Value - (paid.HasValue ? paid.Value : 0));
            U<TextBlock>("SumPrice").Text = Ps.Money(price);
            double tasas = Ps.TasasFor(T, v);
            TextBlock sumIva = U<TextBlock>("SumIva");
            if (price.HasValue && price.Value > 0)
            {
                double b, iva, tx; Ps.SplitIva(price.Value, tasas, out b, out iva, out tx);
                sumIva.Text = tx > 0 ? string.Format(L("sumIvaTasas"), Ps.Money(b), Ps.Money(iva), Ps.Money(tx)) : string.Format(L("sumIva"), Ps.Money(b), Ps.Money(iva));
                sumIva.Visibility = Visibility.Visible;
            }
            else sumIva.Visibility = Visibility.Collapsed;
            // margen (precio - tasas - gastos - comisión) y, en traducciones, el beneficio fijo por documento
            TextBlock sm = U<TextBlock>("SumMargin");
            sm.Foreground = Brush("#667085");
            double? gastos = Ps.Num(Store.Get(v, "gastos"));
            if (T.Id == "traduccion" && price.HasValue && gastos.HasValue)
            {
                double? nd = Ps.Num(Store.Get(v, "ndocs")); if (!nd.HasValue || nd.Value <= 0) nd = 1;
                double want = Config.Cur.TradProfit * nd.Value, got = price.Value - gastos.Value;
                if (Math.Abs(got - want) < 0.005) { sm.Text = string.Format(L("trProfitOk"), Ps.Money(got)); sm.Foreground = Brush("#1E7B34"); }
                else
                {
                    sm.Text = string.Format(L("trProfitBad"), Ps.Money(got), Ps.Money(want), Config.Cur.TradProfit.ToString("0.##", Ps.ES), nd.Value.ToString("0.##", Ps.ES));
                    sm.Foreground = Brush("#B42318");
                }
                sm.Visibility = Visibility.Visible;
            }
            else
            {
                double m, pct, com;
                if (Ps.Margin(T, v, out m, out pct, out com))
                {
                    string p = pct.ToString("0.#", Ps.ES) + " %";
                    sm.Text = com > 0 ? string.Format(L("sumMarginCom"), Ps.Money(m), p, Ps.Money(com)) : string.Format(L("sumMargin"), Ps.Money(m), p);
                    sm.Visibility = Visibility.Visible;
                }
                else sm.Visibility = Visibility.Collapsed;
            }
            U<TextBlock>("SumPaid").Text = Ps.Money(price.HasValue ? (object)(paid.HasValue ? paid.Value : 0) : null);
            U<TextBlock>("SumPend").Text = Ps.Money(pend);
            Brush pb = Brush(pend.HasValue && pend.Value > 0 ? "#B42318" : "#1E7B34");
            U<TextBlock>("SumPend").Foreground = pb; U<TextBlock>("SumPendL").Foreground = pb;
        }

        ClientRow FindDuplicateDoc(string doc)
        {
            if (string.IsNullOrEmpty(doc) || G.Rows == null) return null;
            foreach (ClientRow r in G.Rows)
            {
                if (EditRow != null && r.Row == EditRow.Row) continue;
                if (Ps.EqI((r.Doc ?? "").Trim(), doc)) return r;
            }
            return null;
        }
        void CheckDuplicate()
        {
            ClientRow d = FindDuplicateDoc(((TextBox)ctl["doc"]).Text.Trim());
            if (d != null) { U<TextBlock>("DupWarn").Text = string.Format(L("dupWarn"), d.Id, d.Name); U("DupWarn").Visibility = Visibility.Visible; }
            else U("DupWarn").Visibility = Visibility.Collapsed;
        }

        // ¿alguien cambió este cliente desde que se abrió para editar?
        ClientRec EditConflict()
        {
            if (EditRow == null) return null;
            try
            {
                StoreData s = Store.Read(G.T);
                if (s == null) return null;
                foreach (ClientRec c in s.Clients)
                    if (c.Row == EditRow.Row) return (c.Updated.Length > 0 && c.Updated != (EditRow.Updated ?? "")) ? c : null;
            }
            catch (Exception e) { App.Log("DEBUG conflicto: " + e.Message); }
            return null;
        }
        string[] WhoWhen(ClientRec c)
        {
            string who = c.By;
            if (who.Length == 0 || who == "Excel" || who == "ADMIN") who = L("someoneElse");
            string when = ""; DateTime d;
            if (DateTime.TryParse(c.Updated, Ps.Inv, System.Globalization.DateTimeStyles.None, out d))
                when = d.Date == DateTime.Today ? d.ToString("HH:mm") : d.ToString("dd/MM HH:mm");
            return new string[] { who, when };
        }

        public void SaveFromForm()
        {
            App.Log("Save-FromForm inicio");
            // shortcut: Ctrl+Z de la última acción rápida (UndoInfo) llega en la etapa 3c
            Dictionary<string, object> vals = GetFormValues();
            if (V(vals, "name").Length == 0) { ShowError(L("reqName")); ctl["name"].Focus(); return; }
            string nl = Environment.NewLine + Environment.NewLine;
            ClientRow dup = FindDuplicateDoc(V(vals, "doc"));
            if (dup != null && MessageBox.Show(Win, string.Format(L("dupAsk"), dup.Id, dup.Name, nl), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            ClientRec conf = EditConflict();
            if (conf != null)
            {
                string[] ww = WhoWhen(conf);
                if (MessageBox.Show(Win, string.Format(L("conflictAsk"), ww[0], ww[1]), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            }
            Button save = U<Button>("SaveBtn");
            save.IsEnabled = false; Win.Cursor = Cursors.Wait;
            SetStatus(L("savingMsg"));
            try
            {
                int? row = EditRow != null ? (int?)EditRow.Row : null;
                SaveResult res = Data.SaveClientStore(vals, row);
                bool wasEdit = EditRow != null;
                App.Log(string.Format("{0} cliente {1} · {2} · {3}", wasEdit ? "editó" : "añadió", res.Id, V(vals, "name"), G.T.TitleEs));
                Data.LoadData(false, false);
                SetStatus(string.Format(L(wasEdit ? "savedEdit" : "saved"), res.Id));
                if (!res.Synced) SetStatus(string.Format(L("savedNoExcel"), res.Id));
                ResetForm();
                if (wasEdit) U<ToggleButton>("NavClients").IsChecked = true;
            }
            catch (Exception e)
            {
                App.Log("Save-FromForm: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
                SetStatus("");
            }
            finally { save.IsEnabled = true; Win.Cursor = null; }
        }

        public void EditSelected()
        {
            try
            {
                ClientRow sel = U<DataGrid>("ClientsGrid").SelectedItem as ClientRow;
                if (sel == null) { MessageBox.Show(Win, L("selectFirst"), App.Name, MessageBoxButton.OK, MessageBoxImage.Information); return; }
                U<ToggleButton>("NavNew").IsChecked = true;
                FillFormForEdit(sel);
            }
            catch (Exception e) { App.Log("DEBUG editar: " + e.Message); ShowError(e.Message); }
        }
    }
}

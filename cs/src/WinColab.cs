// Etapa 4e: Colaboradores y abogados (sus trámites, ingresos y comisiones de todos los trámites). Copiado de Get-ColabItems,
// Load-ColabPanel, Update-ColabView, Set-ColabComState, Mark-ColabPaid, Undo-ColabPaid y Open-ColabItem de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace Gestem
{
    public class ColabItem
    {
        public string TramId { get; set; } public string Tram { get; set; } public int Row { get; set; } public string Id { get; set; } public string Name { get; set; } public string Colab { get; set; }
        public DateTime FechaV { get; set; } public string FechaTxt { get; set; } public string Status { get; set; } public string Color { get; set; }
        public double Price { get; set; } public double Paid { get; set; } public double Pend { get; set; } public double Com { get; set; } public bool ComPaid { get; set; }
        public string PriceTxt { get; set; } public string PaidTxt { get; set; } public string PendTxt { get; set; } public string ComTxt { get; set; } public string ComStTxt { get; set; } public string ComColor { get; set; }
    }
    public class ColabListItem { public string Key { get; set; } public string Name { get; set; } public string Glyph { get; set; } public string Info { get; set; } }
    public class ColabPair { public Tramite T; public int Row; public string Id = "", Value = ""; }    // @{ T; Row; Id; Value }

    public partial class MainWin
    {
        public List<ColabItem> ColabItems = new List<ColabItem>();
        bool colabFilling;

        void WireColab()
        {
            U<ToggleButton>("NavColab").Checked += delegate { ShowPanel("PanelColab"); LoadColabPanel(); };
            U<ListBox>("ColabList").SelectionChanged += delegate { UpdateColabView(); };
            U<ToggleButton>("ColabOnlyPend").Click += delegate { UpdateColabView(); };
            U<Button>("ColabPaidBtn").Click += delegate { MarkColabPaid(); };
            U<DataGrid>("ColabGrid").MouseDoubleClick += delegate { OpenColabItem(U<DataGrid>("ColabGrid").SelectedItem as ColabItem); };
        }

        public void SetColabTexts()
        {
            U<TextBlock>("ColabTitle").Text = L("colTitle"); U<TextBlock>("ColabHint").Text = L("colHint");
            U<ContentControl>("ColabPaidBtn").Content = L("colPaidBtn"); U<ContentControl>("ColabOnlyPend").Content = L("colOnlyPend");
            U<TextBlock>("ColK1L").Text = L("colIncome"); U<TextBlock>("ColK2L").Text = L("colCollected"); U<TextBlock>("ColK3L").Text = L("colToCollect");
            U<TextBlock>("ColK4L").Text = L("colCom"); U<TextBlock>("ColK5L").Text = L("colComPaid"); U<TextBlock>("ColK6L").Text = L("colComPend");
            string[] hk = { "colHTram", "colHId", "colHName", "seccolab", "colHDate", "colHStatus", "colHPrice", "colHPaid", "colHPend", "colHCom", "colHComSt" };
            DataGrid g = U<DataGrid>("ColabGrid");
            for (int i = 0; i < hk.Length && i < g.Columns.Count; i++) g.Columns[i].Header = L(hk[i]);
        }

        public static List<ColabItem> GetColabItems()
        {
            List<ColabItem> items = new List<ColabItem>();
            if (G.Cache == null) return items;
            foreach (Tramite t in Tramite.All)
            {
                if (!G.Cache.ContainsKey(t.Id) || t.FieldByKey("colab") == null) continue;
                string title = t.TitleText();
                foreach (ClientRow r in G.Cache[t.Id].Rows)
                {
                    string who = Ps.Str(Store.Get(r.Data, "colab")).Trim().ToUpper();
                    if (who.Length == 0) continue;
                    double price = r.PriceV, pend = r.Pending, paid = Math.Max(0, price - pend);
                    double com = Ps.Commission(r.Data, price);
                    bool paidCom = Ps.Str(Store.Get(r.Data, "compag")) == "Pagada";
                    ColabItem x = new ColabItem();
                    x.TramId = t.Id; x.Tram = title; x.Row = r.Row; x.Id = Ps.Str(r.Id); x.Name = Ps.Str(r.Name); x.Colab = who;
                    x.FechaV = r.Fecha.HasValue ? r.Fecha.Value : DateTime.MinValue;
                    x.FechaTxt = r.Fecha.HasValue ? r.Fecha.Value.ToString("dd/MM/yyyy") : "";
                    x.Status = Ps.Str(r.Status); x.Color = r.Color;
                    x.Price = price; x.Paid = paid; x.Pend = pend; x.Com = com; x.ComPaid = paidCom;
                    x.PriceTxt = price > 0 ? Ps.Money(price) : ""; x.PaidTxt = price > 0 ? Ps.Money(paid) : "";
                    x.PendTxt = pend > 0 ? Ps.Money(pend) : ""; x.ComTxt = com > 0 ? Ps.Money(com) : "";
                    x.ComStTxt = com <= 0 ? "" : paidCom ? L("colStPaid") : L("colStPend");
                    x.ComColor = com <= 0 ? "Transparent" : paidCom ? "#DCFAE6" : "#FEF0C7";
                    items.Add(x);
                }
            }
            return items;
        }

        public void LoadColabPanel()
        {
            ListBox lb = U<ListBox>("ColabList");
            string keep = lb.SelectedItem is ColabListItem ? ((ColabListItem)lb.SelectedItem).Key : null;
            ColabItems = GetColabItems();
            List<ColabListItem> list = new List<ColabListItem>();
            double pendAll = 0; foreach (ColabItem x in ColabItems) if (!x.ComPaid) pendAll += x.Com;
            list.Add(new ColabListItem { Key = "", Name = L("colAll"), Glyph = ((char)0xE716).ToString(), Info = string.Format(L("colItem"), ColabItems.Count, Ps.Money(pendAll)) });
            foreach (string n in CollaboratorNames())
            {
                int cnt = 0; double p = 0;
                foreach (ColabItem x in ColabItems) if (string.Equals(x.Colab, n, StringComparison.CurrentCultureIgnoreCase)) { cnt++; if (!x.ComPaid) p += x.Com; }
                list.Add(new ColabListItem { Key = n, Name = n, Glyph = ((char)0xE77B).ToString(), Info = string.Format(L("colItem"), cnt, Ps.Money(p)) });
            }
            colabFilling = true;
            try
            {
                lb.ItemsSource = list;
                int sel = 0;
                if (!string.IsNullOrEmpty(keep)) for (int i = 0; i < list.Count; i++) if (string.Equals(list[i].Key, keep, StringComparison.CurrentCultureIgnoreCase)) sel = i;
                lb.SelectedIndex = sel;
            }
            finally { colabFilling = false; }
            UpdateColabView();
        }

        public List<ColabItem> UpdateColabView()
        {
            if (colabFilling) return null;
            ColabListItem si = U<ListBox>("ColabList").SelectedItem as ColabListItem;
            string key = si != null ? Ps.Str(si.Key) : "";
            U<TextBlock>("ColabName").Text = key.Length > 0 ? key : L("colTitle");
            List<ColabItem> mine = new List<ColabItem>();
            foreach (ColabItem x in ColabItems) if (key.Length == 0 || string.Equals(x.Colab, key, StringComparison.CurrentCultureIgnoreCase)) mine.Add(x);
            double inc = 0, col = 0, pend = 0, com = 0, comP = 0, comN = 0;
            foreach (ColabItem x in mine) { inc += x.Price; col += x.Paid; pend += x.Pend; com += x.Com; if (x.ComPaid) comP += x.Com; else comN += x.Com; }
            U<TextBlock>("ColK1").Text = Ps.Money(inc); U<TextBlock>("ColK2").Text = Ps.Money(col); U<TextBlock>("ColK3").Text = Ps.Money(pend);
            U<TextBlock>("ColK4").Text = Ps.Money(com); U<TextBlock>("ColK5").Text = Ps.Money(comP); U<TextBlock>("ColK6").Text = Ps.Money(comN);
            List<ColabItem> show = mine;
            if (U<ToggleButton>("ColabOnlyPend").IsChecked == true) { show = new List<ColabItem>(); foreach (ColabItem x in mine) if (x.Com > 0 && !x.ComPaid) show.Add(x); }
            show = new List<ColabItem>(show);
            Ps.Sort(show, delegate(ColabItem a, ColabItem b) { return b.FechaV.CompareTo(a.FechaV); });    // Sort-Object FechaV -Descending (empates como PS)
            DataGrid g = U<DataGrid>("ColabGrid");
            g.ItemsSource = show;
            U<TextBlock>("ColabEmpty").Text = L("colNone");
            U("ColabEmpty").Visibility = ColabItems.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            g.Visibility = ColabItems.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
            return show;
        }

        // cambia «Pago comisión» directamente en los datos de GESTEM (no va al Excel); devuelve lo que había, para deshacer
        public static List<ColabPair> SetColabComState(List<ColabPair> pairs)
        {
            List<ColabPair> done = new List<ColabPair>();
            List<string> order = new List<string>(); Dictionary<string, List<ColabPair>> groups = new Dictionary<string, List<ColabPair>>(StringComparer.CurrentCultureIgnoreCase);    // Group-Object (orden de aparición)
            foreach (ColabPair p in pairs) { if (!groups.ContainsKey(p.T.Id)) { groups[p.T.Id] = new List<ColabPair>(); order.Add(p.T.Id); } groups[p.T.Id].Add(p); }
            foreach (string tid in order)
            {
                Tramite t = groups[tid][0].T;
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    if (s == null) continue;
                    foreach (ColabPair p in groups[tid])
                    {
                        ClientRec c = null;
                        if (p.Row > 0) foreach (ClientRec x in s.Clients) if (x.Row == p.Row) { c = x; break; }
                        if (c == null && !string.IsNullOrEmpty(p.Id)) foreach (ClientRec x in s.Clients) if (Ps.Str(x.Id) == p.Id) { c = x; break; }
                        if (c == null) continue;
                        done.Add(new ColabPair { T = t, Row = c.Row, Id = Ps.Str(c.Id), Value = Ps.Str(Store.Get(c.F, "compag")) });
                        c.F["compag"] = string.IsNullOrEmpty(p.Value) ? null : p.Value;
                        c.Updated = App.Now().ToString("s", Ps.Inv); c.By = App.User.Length > 0 ? App.User : "ADMIN";
                    }
                    Store.Write(t, s);
                }
                finally { Store.ExitLock(lk); }
            }
            return done;
        }

        public void MarkColabPaid()
        {
            List<ColabItem> sel = new List<ColabItem>();
            foreach (object o in U<DataGrid>("ColabGrid").SelectedItems) { ColabItem x = o as ColabItem; if (x != null && x.Com > 0 && !x.ComPaid) sel.Add(x); }
            if (sel.Count == 0) { Info(L("colSelFirst")); return; }
            Busy = true; Win.Cursor = Cursors.Wait;
            SetStatus(L("savingMsg"));
            try
            {
                List<ColabPair> pairs = new List<ColabPair>();
                foreach (ColabItem x in sel) pairs.Add(new ColabPair { T = Tramite.ById(x.TramId), Row = x.Row, Id = x.Id, Value = "Pagada" });
                List<ColabPair> old = SetColabComState(pairs);
                List<string> who = new List<string>(); foreach (ColabItem x in sel) if (!who.Contains(x.Colab)) who.Add(x.Colab);    // Select-Object -Unique
                App.Log(string.Format("comisión pagada: {0} trámite(s) · {1}", old.Count, string.Join(", ", who)));
                Data.LoadData(true, true);
                LoadColabPanel();
                UndoInfo u = new UndoInfo(); u.Kind = "colab"; u.Colab = old; Undo = u;
                Toast(string.Format(L("colPaidDone"), old.Count), true, 15);
            }
            catch (Exception e)
            {
                App.Log("comisión pagada: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Busy = false; Win.Cursor = null; SetStatus(""); }
        }

        void UndoColabPaid(UndoInfo u)
        {
            Busy = true; Win.Cursor = Cursors.Wait;
            try
            {
                List<ColabPair> r = SetColabComState(u.Colab);
                App.Log(string.Format("deshizo comisión pagada: {0} trámite(s)", r.Count));
                Data.LoadData(true, true);
                if (U("PanelColab").Visibility == Visibility.Visible) LoadColabPanel();
                Toast(L("cmUndone"), false, 7);
            }
            catch (Exception e)
            {
                App.Log("deshacer comisión: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Busy = false; Win.Cursor = null; }
        }

        // doble clic: abre la ficha del cliente en su trámite
        public void OpenColabItem(ColabItem x)
        {
            if (x == null) return;
            int i = Tramite.All.IndexOf(Tramite.ById(x.TramId));
            if (i < 0) return;
            if (G.T.Id != x.TramId) U<ComboBox>("TramiteBox").SelectedIndex = i;
            if (G.T.Id != x.TramId) return;
            ClientRow r = RowByNumber(x.Row);
            if (r == null) { Toast(L("cmUndoFail"), false, 7); return; }
            U<ToggleButton>("NavNew").IsChecked = true;
            FillFormForEdit(r);
        }
    }
}

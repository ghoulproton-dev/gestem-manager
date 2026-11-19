// Etapa 3c: menú del clic derecho en la lista, atajos (estado, pagado, responsable), eliminar cliente y Ctrl+Z.
// Copiado de Set-CmHeaders, Fill-CmMenu, Quick-SetField, Set-RespBulk, Undo-RespBulk, Remove-Client, Undo-RemoveClient,
// Invoke-QuickUndo y Copy-ClientText de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Gestem
{
    // $script:UndoInfo: lo necesario para deshacer la última acción rápida
    public class UndoInfo
    {
        public string Kind = "";                 // '' (un dato) | resp | colab | delete
        public string TramiteId = ""; public int Row; public string Key = ""; public object Old, OldPaid;   // un dato
        public Tramite T; public List<RespUndo> Items = new List<RespUndo>(); public string Who = "";       // resp
        public bool Hand; public string Name = "", Id = "", Uid;                                           // delete
        public List<ColabPair> Colab;                                                                      // colab
    }
    public class RespUndo { public int Row; public string Id = "", Old = ""; }

    public partial class MainWin
    {
        public UndoInfo Undo;
        public bool Busy;
        public Action<string, string> ToastHook;    // pruebas: qué aviso se mostró
        public Action<string> ErrorHook;            // pruebas: qué error se mostró
        MenuItem cmFolder, cmStatus, cmPaid, cmResp, cmCopyName, cmCopyDoc, cmCopyTel, cmExcel;

        MenuItem NewCmItem(int glyph)
        {
            MenuItem m = new MenuItem(); m.Icon = CmIcon(glyph); return m;
        }
        TextBlock CmIcon(int glyph)
        {
            TextBlock tb = new TextBlock(); tb.Text = ((char)glyph).ToString(); tb.FontFamily = new FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets"); tb.FontSize = 13;
            tb.Foreground = Brush("#9E1820"); tb.HorizontalAlignment = HorizontalAlignment.Center; tb.VerticalAlignment = VerticalAlignment.Center;
            return tb;
        }

        void WireMenu()
        {
            DataGrid g = U<DataGrid>("ClientsGrid");
            ContextMenu cm = g.ContextMenu;
            U<MenuItem>("CmEdit").Icon = CmIcon(0xE70F); U<MenuItem>("CmReceipt").Icon = CmIcon(0xE8A5);
            U<MenuItem>("CmDocs").Icon = CmIcon(0xE723); U<MenuItem>("CmDelete").Icon = CmIcon(0xE74D);
            cmFolder = NewCmItem(0xE838); cmStatus = NewCmItem(0xE7C1); cmPaid = NewCmItem(0xE8C7); cmCopyName = NewCmItem(0xE8C8);
            cmCopyDoc = NewCmItem(0xE8C8); cmCopyTel = NewCmItem(0xE717); cmExcel = NewCmItem(0xE8A7); cmResp = NewCmItem(0xE77B);
            foreach (Control it in new Control[] { cmFolder, new Separator(), cmStatus, cmPaid, cmResp, new Separator(), cmCopyName, cmCopyDoc, cmCopyTel, new Separator(), cmExcel }) cm.Items.Add(it);
            // shortcut: recibo (etapa 5) y abrir el Excel (etapa 4) aún no están: se ven desactivados
            foreach (MenuItem it in new MenuItem[] { U<MenuItem>("CmReceipt"), cmExcel }) it.IsEnabled = false;

            U<MenuItem>("CmEdit").Click += delegate { EditSelected(); };
            U<MenuItem>("CmDelete").Click += delegate { ClientRow r = SelectedClient(); if (r != null) RemoveClient(r, true); };
            cmPaid.Click += delegate { QuickSetField("paystate", "Pagado", false); };
            cmCopyName.Click += delegate { ClientRow r = SelectedClient(); if (r != null) CopyClientText(r.Name); };
            cmCopyDoc.Click += delegate { ClientRow r = SelectedClient(); if (r != null) CopyClientText(r.Doc); };
            cmCopyTel.Click += delegate { ClientRow r = SelectedClient(); if (r != null) CopyClientText(Ps.Str(Store.Get(r.Data, "tel"))); };
            g.ContextMenuOpening += delegate(object s, ContextMenuEventArgs e)
            {
                ClientRow sel = g.SelectedItem as ClientRow;
                if (sel == null) { e.Handled = true; return; }    // sin cliente debajo: no hay menú
                try { FillCmMenu(sel); } catch (Exception x) { App.Log("menú: " + x.Message); }
            };
            // clic derecho: selecciona la fila que hay debajo antes de abrir el menú
            g.PreviewMouseRightButtonDown += delegate(object s, MouseButtonEventArgs e)
            {
                DependencyObject o = e.OriginalSource as DependencyObject;
                while (o != null && !(o is DataGridRow)) o = (o is Visual || o is System.Windows.Media.Media3D.Visual3D) ? VisualTreeHelper.GetParent(o) : LogicalTreeHelper.GetParent(o);
                DataGridRow row = o as DataGridRow;
                if (row != null && !row.IsSelected) { g.SelectedItems.Clear(); row.IsSelected = true; }   // si ya estaba seleccionada, se mantiene la selección
            };
        }

        public void SetCmHeaders()
        {
            if (cmFolder == null) return;
            cmFolder.Header = L("cmFolder"); cmStatus.Header = L("cmStatus"); cmPaid.Header = L("cmPaid"); cmResp.Header = L("cmResp");
            cmCopyName.Header = L("cmCopyName"); cmCopyDoc.Header = L("cmCopyDoc"); cmCopyTel.Header = L("cmCopyTel"); cmExcel.Header = L("cmExcel");
        }

        ClientRow SelectedClient()
        {
            ClientRow sel = U<DataGrid>("ClientsGrid").SelectedItem as ClientRow;
            if (sel == null && !App.Quiet) MessageBox.Show(Win, L("selectFirst"), App.Name, MessageBoxButton.OK, MessageBoxImage.Information);
            return sel;
        }

        // el menú se ajusta al cliente: estados del trámite, y solo lo que tiene sentido
        public void FillCmMenu(ClientRow sel)
        {
            Tramite T = G.T;
            cmStatus.Items.Clear();
            Field f = T.FieldByKey("status");
            if (f != null)
                foreach (string s in GetList(f.List))
                {
                    if (Ps.Str(s).Trim().Length == 0) continue;
                    MenuItem it = new MenuItem(); it.Header = s.Replace("_", "__"); it.Tag = s;
                    if (Ps.Str(sel.Status) == s) { it.IsChecked = true; it.FontWeight = FontWeights.SemiBold; }
                    it.Click += delegate(object o, RoutedEventArgs e) { QuickSetField("status", (string)((MenuItem)o).Tag, false); };
                    cmStatus.Items.Add(it);
                }
            cmStatus.IsEnabled = cmStatus.Items.Count > 0;
            // asignar responsable: a todos los clientes seleccionados (Ctrl+A, Mayús o Ctrl + clic)
            int nSel = U<DataGrid>("ClientsGrid").SelectedItems.Count;
            cmResp.Header = nSel > 1 ? string.Format(L("cmRespN"), nSel) : L("cmResp");
            cmResp.Items.Clear();
            List<string> people = new List<string>();
            List<string> src = new List<string>(Config.Cur.Colleagues);
            string[] lc; if (G.Lists != null && G.Lists.TryGetValue("colleagues", out lc) && lc != null) src.AddRange(lc);
            foreach (string p0 in src) { string p = Ps.Str(p0).Trim().ToUpper(); if (p.Length > 0 && !PsH.ContainsI(people, p)) people.Add(p); }
            foreach (string p in people)
            {
                MenuItem it = new MenuItem(); it.Header = p.Replace("_", "__"); it.Tag = p;
                if (nSel == 1 && string.Equals(Ps.Str(sel.Resp).Trim(), p, StringComparison.OrdinalIgnoreCase)) { it.IsChecked = true; it.FontWeight = FontWeights.SemiBold; }
                it.Click += delegate(object o, RoutedEventArgs e) { SetRespBulk((string)((MenuItem)o).Tag, true); };
                cmResp.Items.Add(it);
            }
            cmResp.IsEnabled = people.Count > 0 && T.FieldByKey("resp") != null;
            cmPaid.IsEnabled = T.FieldByKey("paystate") != null && Ps.Str(sel.Pay) != "Pagado";
            cmCopyDoc.IsEnabled = Ps.Str(sel.Doc).Trim().Length > 0;
            cmCopyTel.IsEnabled = Ps.Str(Store.Get(sel.Data, "tel")).Trim().Length > 0;
        }
        // para la prueba: lo que muestra el menú
        public string CmMenuText()
        {
            List<string> st = new List<string>(); foreach (MenuItem it in cmStatus.Items) st.Add(Ps.Str(it.Tag) + (it.IsChecked ? "*" : ""));
            List<string> rs = new List<string>(); foreach (MenuItem it in cmResp.Items) rs.Add(Ps.Str(it.Tag) + (it.IsChecked ? "*" : ""));
            return "st=" + string.Join(",", st.ToArray()) + "|" + cmStatus.IsEnabled + " resp=" + Ps.Str(cmResp.Header) + "|" + string.Join(",", rs.ToArray()) + "|" + cmResp.IsEnabled
                + " paid=" + cmPaid.IsEnabled + " doc=" + cmCopyDoc.IsEnabled + " tel=" + cmCopyTel.IsEnabled;
        }

        void Toast(string text, bool undo, int seconds)
        {
            if (undo) ShowToast(text, L("cmUndo"), delegate { InvokeQuickUndo(); }, seconds); else ShowToast(text, "", null, 7);
        }
        static string Who() { return App.User.Length > 0 ? App.User : "ADMIN"; }

        // cambia un solo dato sin abrir la ficha: pasa por el guardado normal (GESTEM + registro)
        public void QuickSetField(string key, object value, bool noUndo)
        {
            ClientRow sel = U<DataGrid>("ClientsGrid").SelectedItem as ClientRow;
            if (sel == null) return;
            Field f = G.T.FieldByKey(key);
            if (f == null) return;
            object old = Store.Get(sel.Data, key), oldPaid = Store.Get(sel.Data, "paid");
            int rowNo = sel.Row; string tid = G.T.Id;
            bool ok = false;
            try
            {
                FillFormForEdit(sel);
                filling = true;
                try { SetFieldValue(f, value); } finally { filling = false; }
                UpdatePreview();
                SaveFromForm();
                ok = EditRow == null;    // SaveFromForm limpia la ficha solo si guardó
            }
            catch (Exception e)
            {
                App.Log("atajo (" + key + "): " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { if (EditRow != null) { try { ResetForm(); } catch { } } }
            if (ok && !noUndo)
            {
                UndoInfo u = new UndoInfo(); u.TramiteId = tid; u.Row = rowNo; u.Key = key; u.Old = old; u.OldPaid = oldPaid;
                Undo = u;
                Toast(string.Format(L("cmDone"), f.LabelText(), Ps.Str(value)), true, 8);
            }
        }

        // pone el mismo responsable a todos los clientes seleccionados, de una vez
        public void SetRespBulk(string who, bool ask)
        {
            List<ClientRow> sel = new List<ClientRow>(); foreach (object o in U<DataGrid>("ClientsGrid").SelectedItems) { ClientRow r = o as ClientRow; if (r != null) sel.Add(r); }
            if (sel.Count == 0 || Ps.Str(who).Trim().Length == 0) return;
            Tramite t = G.T;
            if (t.FieldByKey("resp") == null) return;
            if (sel.Count > 1 && ask && MessageBox.Show(Win, string.Format(L("respAsk"), sel.Count, who), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            int n = 0;
            Busy = true; Win.Cursor = Cursors.Wait;
            SetStatus(L("savingMsg"));
            try
            {
                List<RespUndo> undo = new List<RespUndo>();   // para poder deshacer: responsable que tenía cada uno
                List<ClientRow> norm = new List<ClientRow>(), hand = new List<ClientRow>();
                foreach (ClientRow r in sel) if (r.Row < Hand.RowBase) norm.Add(r); else hand.Add(r);
                if (norm.Count > 0)
                {
                    string lk = Store.EnterLock();
                    try
                    {
                        StoreData s = Store.Read(t);
                        foreach (ClientRow r in norm)
                        {
                            ClientRec c = null; foreach (ClientRec x in s.Clients) if (x.Row == r.Row) { c = x; break; }
                            if (c == null) continue;
                            RespUndo ru = new RespUndo(); ru.Row = c.Row; ru.Id = c.Id; ru.Old = Ps.Str(Store.Get(c.F, "resp")); undo.Add(ru);
                            c.F["resp"] = who; c.Sync = false; c.Updated = App.Now().ToString("s", Ps.Inv); c.By = Who(); n++;
                        }
                        Store.Write(t, s);
                    }
                    finally { Store.ExitLock(lk); }
                }
                // clientes de una pestaña hecha a mano: pasan a GESTEM con el responsable puesto
                foreach (ClientRow r in hand)
                {
                    Dictionary<string, object> vals = new Dictionary<string, object>(); foreach (KeyValuePair<string, object> kv in r.Data) vals[kv.Key] = kv.Value;
                    string oldResp = Ps.Str(Store.Get(vals, "resp"));
                    vals["resp"] = who;
                    SaveResult res = Data.SaveClientStore(vals, r.Row); n++;
                    if (res != null && !string.IsNullOrEmpty(res.Id)) { RespUndo ru = new RespUndo(); ru.Row = 0; ru.Id = res.Id; ru.Old = oldResp; undo.Add(ru); }
                }
                Data.SyncPendingToExcel();
                App.Log(string.Format("asignó {0} cliente(s) a {1} · {2}", n, who, t.TitleEs));
                ReloadList();
                UndoInfo u = new UndoInfo(); u.Kind = "resp"; u.T = t; u.Items = undo; u.Who = who; Undo = u;
                Toast(string.Format(L("respDone"), n, who), true, 15);
            }
            catch (Exception e)
            {
                App.Log("asignar responsable: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Busy = false; Win.Cursor = null; SetStatus(""); }
        }

        // vuelve a poner el responsable que tenía cada cliente antes de «Asignar responsable»
        void UndoRespBulk(UndoInfo u)
        {
            Tramite t = u.T; int n = 0;
            Busy = true; Win.Cursor = Cursors.Wait;
            SetStatus(L("savingMsg"));
            try
            {
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    foreach (RespUndo it in u.Items)
                    {
                        ClientRec c = null;
                        if (it.Row > 0) foreach (ClientRec x in s.Clients) if (x.Row == it.Row) { c = x; break; }
                        if (c == null && it.Id.Length > 0) foreach (ClientRec x in s.Clients) if (x.Id == it.Id) { c = x; break; }
                        if (c == null) continue;
                        c.F["resp"] = it.Old; c.Sync = false; c.Updated = App.Now().ToString("s", Ps.Inv); c.By = Who(); n++;
                    }
                    Store.Write(t, s);
                }
                finally { Store.ExitLock(lk); }
                Data.SyncPendingToExcel();
                App.Log(string.Format("deshizo asignar responsable ({0}): {1} cliente(s) · {2}", u.Who, n, t.TitleEs));
                ReloadList();
                Toast(L("cmUndone"), false, 7);
            }
            catch (Exception e)
            {
                App.Log("deshacer responsable: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Busy = false; Win.Cursor = null; SetStatus(""); }
        }

        public void RemoveClient(ClientRow r, bool ask)
        {
            Tramite t = G.T;
            string name = Ps.Str(r.Name), id = Ps.Str(r.Id).Length > 0 ? r.Id : "—";
            if (ask && MessageBox.Show(Win, string.Format(L("delAsk"), name, id), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
            try
            {
                UndoInfo u = new UndoInfo(); u.Kind = "delete"; u.T = t; u.Hand = r.Row >= Hand.RowBase; u.Name = name; u.Id = Ps.Str(r.Id);
                if (!u.Hand)
                {
                    StoreData s = Store.Read(t);
                    ClientRec c = null; foreach (ClientRec x in s.Clients) if (x.Row == r.Row) { c = x; break; }
                    if (c == null) throw new GestemException(L("cmUndoFail"));
                    u.Uid = c.Uid;
                }
                Data.SetClientDeleted(t, u, true);
                App.Log(string.Format("eliminó el cliente {0} {1} · {2}", id, name, t.TitleEs));
                ReloadList();
                Undo = u;
                Toast(string.Format(L("delDone"), name), true, 15);
            }
            catch (Exception e)
            {
                App.Log("eliminar cliente: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
        }
        void UndoRemoveClient(UndoInfo u)
        {
            try
            {
                Data.SetClientDeleted(u.T, u, false);
                App.Log(string.Format("deshizo eliminar el cliente {0} {1} · {2}", u.Id, u.Name, u.T.TitleEs));
                ReloadList();
                Toast(L("cmUndone"), false, 7);
            }
            catch (Exception e)
            {
                App.Log("deshacer eliminar: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
        }

        public void InvokeQuickUndo()
        {
            UndoInfo u = Undo; Undo = null;
            if (u == null) { Toast(L("undoNone"), false, 7); return; }
            if (u.Kind == "resp") { UndoRespBulk(u); return; }
            if (u.Kind == "colab") { UndoColabPaid(u); return; }
            if (u.Kind == "delete") { UndoRemoveClient(u); return; }
            if (G.T.Id != u.TramiteId) { Undo = u; Toast(L("undoOther"), false, 7); return; }
            ClientRow r = RowByNumber(u.Row);
            if (r == null) { Toast(L("cmUndoFail"), false, 7); return; }
            DataGrid g = U<DataGrid>("ClientsGrid");
            g.SelectedItem = r;
            QuickSetField(u.Key, u.Old, true);
            if (u.Key == "paystate" && u.OldPaid != null)
            {
                ClientRow r2 = RowByNumber(u.Row);
                if (r2 != null) { g.SelectedItem = r2; QuickSetField("paid", u.OldPaid, true); }
            }
            Toast(L("cmUndone"), false, 7);
        }
        static ClientRow RowByNumber(int row) { if (G.Rows != null) foreach (ClientRow r in G.Rows) if (r.Row == row) return r; return null; }

        void ReloadList() { Data.LoadData(false, false); }    // Load-Data: también rellena la lista

        public void CopyClientText(string text)
        {
            text = Ps.Str(text).Trim();
            if (text.Length == 0) return;
            try { Clipboard.SetText(text); SetStatus(string.Format(L("cmCopied"), text)); } catch (Exception e) { App.Log("copiar: " + e.Message); }
        }
    }
}

// Prueba de igualdad con la versión PowerShell: lee los datos de una carpeta de prueba, calcula cada fila
// y vuelve a guardar los archivos. tools\pstest.ps1 hace lo mismo con las funciones de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Text;

namespace Gestem
{
    public static class SelfTest
    {
        public static string N(double v) { return v.ToString("0.##", Ps.Inv); }

        static string WR(FileResult r) { return r == null ? "null" : r.Row + ":" + r.Id; }
        static string WRL(List<FileResult> l) { if (l == null) return "null"; List<string> p = new List<string>(); foreach (FileResult r in l) p.Add(WR(r)); return string.Join(",", p.ToArray()); }

        public static int Run(string testDir)
        {
            StringBuilder o = new StringBuilder();
            App.AddSafeRoot(testDir);
            App.DataDir = testDir; App.BackupDir = Path.Combine(testDir, "copias"); App.LogPath = Path.Combine(testDir, "cs_log.txt"); App.CfgPath = Path.Combine(testDir, "config_cs.json"); FbState.Path = Path.Combine(testDir, "cs_firebase.json");
            Directory.CreateDirectory(App.BackupDir);
            Config.Cur = new Config();
            Config.Cur.ExcelPath = Path.Combine(testDir, "cs", "x.xlsx");
            // copia de los datos tal como llegan, para comprobar al final que al abrir no se pierde ningún cliente con uid (test\dups)
            string dd = Path.Combine(testDir, "dups"), dsrc = Path.GetDirectoryName(Config.Cur.ExcelPath);
            Directory.CreateDirectory(Path.Combine(dd, "GESTEM datos"));
            File.Copy(Config.Cur.ExcelPath, Path.Combine(dd, "x.xlsx"), true);
            foreach (string f in Directory.GetFiles(Path.Combine(dsrc, "GESTEM datos"), "*.json")) File.Copy(f, Path.Combine(dd, "GESTEM datos", Path.GetFileName(f)), true);
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = Store.Read(t);
                if (s == null) { o.Append("T " + t.Id + " (sin datos)\r\n"); continue; }
                TramiteData d = Store.BuildData(t, s);
                double pend = 0, price = 0, paid = 0, net = 0;
                foreach (ClientRow r in d.Rows) { pend += r.Pending; price += r.PriceV; paid += r.PaidV; net += r.NetV; }
                o.Append("T " + t.Id + " rows=" + d.Rows.Count + " firstEmpty=" + (d.FirstEmpty.HasValue ? d.FirstEmpty.Value.ToString() : "") + " nextId=" + d.NextId
                    + " pend=" + N(pend) + " price=" + N(price) + " paid=" + N(paid) + " net=" + N(net) + "\r\n");
                foreach (ClientRow r in d.Rows)
                    o.Append("R " + t.Id + " " + r.Row + " " + r.Id + "|" + r.Name + "|" + r.Doc + "|" + r.Tipo + "|" + r.Status + "|" + r.Color + "|" + r.Pay + "|"
                        + N(r.Pending) + "|" + r.PendingText + "|" + Ps.Str(r.Days) + "|" + r.Loc + "|" + r.Resp + "|" + N(r.PriceV) + "|" + N(r.PaidV) + "|" + N(r.NetV) + "|"
                        + (r.Fecha.HasValue ? r.Fecha.Value.ToString("yyyy-MM-dd", Ps.Inv) : "") + "\r\n");
                List<string> keys = new List<string>(d.Lists.Keys); keys.Sort(StringComparer.Ordinal);
                foreach (string k in keys) o.Append("L " + t.Id + " " + k + "=" + string.Join("|", d.Lists[k]) + "\r\n");
                // resolver un registro como al guardar
                foreach (ClientRec c in s.Clients)
                {
                    Dictionary<string, object> fv = new Dictionary<string, object>(c.F);
                    Store.Resolve(t, fv, s.Meta);
                    o.Append("V " + t.Id + " " + c.Row + " price=" + Ps.Str(fv.ContainsKey("price") ? fv["price"] : null) + " paid=" + Ps.Str(fv.ContainsKey("paid") ? fv["paid"] : null) + "\r\n");
                }
                Store.Write(t, s);
            }
            // etapa 2: listas desplegables y pestañas hechas a mano, leídas del archivo
            string xl = Config.Cur.ExcelPath;
            System.Collections.Hashtable dv = Xlsx.ReadValidationLists(xl);
            List<string> dk = new List<string>(); foreach (object k in dv.Keys) dk.Add((string)k); dk.Sort(StringComparer.Ordinal);
            foreach (string k in dk)
            {
                DvSheet sh = (DvSheet)dv[k];
                List<string> ck = new List<string>(); foreach (object c in sh.Cols.Keys) ck.Add((string)c); ck.Sort(StringComparer.Ordinal);
                foreach (string c in ck) { object v = sh.Cols[c]; o.Append("D " + k + " [" + sh.Name + "] " + c + "=" + (v is string ? (string)v : string.Join("|", (string[])v)) + "\r\n"); }
            }
            foreach (Tramite t in Tramite.All) o.Append("U " + t.Id + " " + Hand.UpdateListsFromFile(t, dv) + "\r\n");
            o.Append("S changed=" + Hand.SyncTabs(true) + "\r\n");
            List<string> ht = new List<string>(); foreach (object k in Hand.HandTabs.Keys) ht.Add((string)k); ht.Sort(StringComparer.Ordinal);
            foreach (string k in ht) o.Append("HT " + k + "=" + Hand.HandTabs[k] + "\r\n");
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = Store.Read(t);
                if (s == null) { o.Append("T2 " + t.Id + " (sin datos)\r\n"); continue; }
                TramiteData d = Store.BuildData(t, s);
                o.Append("T2 " + t.Id + " rows=" + d.Rows.Count + "\r\n");
                List<ClientRec> sc = new List<ClientRec>(s.Clients); Store.StableSortByRow(sc);
                for (int i = 0; i < sc.Count; i++)
                {
                    ClientRow r = d.Rows[i];
                    StringBuilder line = new StringBuilder("R2 " + t.Id + " " + r.Row + " " + r.Id);
                    foreach (Field f in t.Fields) line.Append("|" + f.Key + "=" + Ps.Str(Store.ToStoreValue(f, Store.Get(sc[i].F, f.Key))));
                    o.Append(line.ToString() + "\r\n");
                }
                List<string> keys = new List<string>(d.Lists.Keys); keys.Sort(StringComparer.Ordinal);
                foreach (string k in keys) o.Append("L2 " + t.Id + " " + k + "=" + string.Join("|", d.Lists[k]) + "\r\n");
            }
            // etapa 2b: escribir clientes directamente en el archivo
            Config.Cur.Backup = false;
            Config.Cur.ExcelRepaired = false;
            ExcelFile.RepairExcelFile();
            o.Append("RX repaired=" + Config.Cur.ExcelRepaired + "\r\n");
            {
                Tramite t = Tramite.ById("canje");
                StoreData s = Store.Read(t);
                TramiteData d = Store.BuildData(t, s);
                Dictionary<int, ClientRec> byRow = new Dictionary<int, ClientRec>();
                foreach (ClientRec c in s.Clients) if (!byRow.ContainsKey(c.Row)) byRow[c.Row] = c;
                Dictionary<string, object> v1 = new Dictionary<string, object>(byRow[13].F);
                v1["status"] = "Finalizado"; v1["paystate"] = "Parcial"; v1["paid"] = 50.0;
                o.Append("W1 " + WR(ExcelFile.SaveOne(t, ExcelFile.ClientOps(t, v1, d.Tariff), 13, Ps.Str(v1["name"]), false)) + "\r\n");
                Dictionary<string, object> v2 = new Dictionary<string, object>();
                foreach (Field f in t.Fields) v2[f.Key] = null;
                v2["name"] = "PRUEBA GESTEM CS"; v2["doc"] = "X1234567Z"; v2["tipo"] = "B (Coche)"; v2["status"] = "Documentación";
                v2["paystate"] = "Pendiente"; v2["fsol"] = new DateTime(2026, 11, 14); v2["obs"] = " espacio ";
                Dictionary<string, object> v3 = new Dictionary<string, object>(byRow[20].F);
                v3["obs"] = "nota de prueba"; v3["price"] = 99.0;
                List<FileItem> batch = new List<FileItem>();
                FileItem i2 = new FileItem(); i2.Ops = ExcelFile.ClientOps(t, v2, d.Tariff); i2.Row = d.FirstEmpty; i2.Expect = "PRUEBA GESTEM CS"; i2.IsNew = true; batch.Add(i2);
                FileItem i3 = new FileItem(); i3.Ops = ExcelFile.ClientOps(t, v3, d.Tariff); i3.Row = 20; i3.Expect = Ps.Str(v3["name"]); i3.IsNew = false; batch.Add(i3);
                o.Append("W2 " + WRL(ExcelFile.SaveMany(t, batch, false)) + "\r\n");
                List<FileItem> allIt = new List<FileItem>();
                foreach (ClientRec c in s.Clients) { FileItem fi = new FileItem(); fi.Ops = ExcelFile.ClientOps(t, c.F, d.Tariff); fi.Row = c.Row; fi.Expect = Ps.Str(c.F["name"]); allIt.Add(fi); }
                o.Append("W3 " + WRL(ExcelFile.SaveMany(t, allIt, true)) + " fixed=" + ExcelFile.SumFixed + "\r\n");
                Tramite ta = Tramite.ById("autonomo");
                StoreData sa = Store.Read(ta); TramiteData da = Store.BuildData(ta, sa);
                o.Append("W4 " + WR(ExcelFile.SaveOne(ta, ExcelFile.ClientOps(ta, sa.Clients[0].F, da.Tariff), sa.Clients[0].Row, Ps.Str(sa.Clients[0].F["name"]), false)) + "\r\n");
            }
            // etapa 2c: Excel por COM (leer la pestaña, guardar con Excel abierto, crear una pestaña nueva)
            Book h = null;
            try
            {
                G.T = Tramite.ById("canje");
                h = ExcelCom.Open(false);
                TramiteData d = ExcelCom.ReadTramiteData(h);
                o.Append("C1 rows=" + d.Rows.Count + " firstEmpty=" + d.FirstEmpty + " nextId=" + d.NextId + " def=" + Ps.Str(d.DefaultPrice) + "\r\n");
                foreach (ClientRow r in d.Rows)
                    o.Append("CR " + r.Row + " " + r.Id + "|" + r.Name + "|" + r.Doc + "|" + r.Tipo + "|" + r.Status + "|" + r.Pay + "|" + N(r.Pending) + "|" + Ps.Str(r.Days) + "|" + r.Resp + "|" + N(r.PriceV) + "|" + N(r.PaidV) + "|" + N(r.NetV) + "|" + (r.Fecha.HasValue ? r.Fecha.Value.ToString("yyyy-MM-dd", Ps.Inv) : "") + "\r\n");
                List<string> lk2 = new List<string>(d.Lists.Keys); lk2.Sort(StringComparer.Ordinal);
                foreach (string k in lk2) o.Append("CL " + k + "=" + string.Join("|", d.Lists[k]) + "\r\n");
                List<string> tk = new List<string>(d.Tariff.Keys); tk.Sort(StringComparer.Ordinal);
                foreach (string k in tk) o.Append("CT " + k + "=" + N(d.Tariff[k]) + "\r\n");
                ExcelCom.Close(h, false); h = null;
                // guardar con Excel (como cuando el archivo está abierto)
                StoreData s = Store.Read(G.T); TramiteData sd = Store.BuildData(G.T, s); G.Tariff = sd.Tariff;
                ClientRec c13 = null; foreach (ClientRec c in s.Clients) if (c.Row == 14) { c13 = c; break; }
                Dictionary<string, object> v = new Dictionary<string, object>(c13.F);
                v["status"] = "Finalizado"; v["paystate"] = "Pagado"; v["obs"] = "guardado con Excel"; v["fres"] = new DateTime(2026, 11, 13);
                h = ExcelCom.Open(true);
                FileResult fr = ExcelCom.SaveClient(v, 14, Ps.Str(v["name"]), false, h, false);
                Dictionary<string, object> vn = new Dictionary<string, object>();
                foreach (Field f in G.T.Fields) vn[f.Key] = null;
                vn["name"] = "NUEVO CON EXCEL"; vn["tipo"] = "B (Coche)"; vn["status"] = "Documentación"; vn["paystate"] = "Pendiente";
                FileResult fn = ExcelCom.SaveClient(vn, null, "NUEVO CON EXCEL", true, h, false);
                G.T = Tramite.ById("autonomo");
                StoreData sa = Store.Read(G.T); G.Tariff = sa.Meta.Tariff;
                FileResult fa = ExcelCom.SaveClient(new Dictionary<string, object>(sa.Clients[0].F), null, Ps.Str(sa.Clients[0].F["name"]), true, h, false);
                ExcelCom.Close(h, true); h = null;
                o.Append("C2 " + WR(fr) + " " + WR(fn) + " " + WR(fa) + "\r\n");
            }
            catch (Exception e) { o.Append("C! " + e.Message + "\r\n"); File.WriteAllText(Path.Combine(testDir, "cs_com_error.txt"), e.ToString()); if (h != null) { try { ExcelCom.Close(h, false); } catch { } } }
            // etapa 2d: huella para Firebase (pendientes), cliente eliminado, avisos de requerimientos
            {
                Tramite t = Tramite.ById("sis");
                StoreData s = Store.Read(t);
                s.Clients[0].F["obs"] = "cambio 2d";
                s.Clients[1].Del = new DelMark(); s.Clients[1].Del.By = "PRUEBA"; s.Clients[1].Del.At = "2026-11-19T10:00:00";
                Store.Write(t, s);
                s = Store.Read(t); TramiteData d = Store.BuildData(t, s);
                int del = 0; foreach (ClientRec c in s.Clients) if (c.Del != null) del++;
                o.Append("FB rows=" + d.Rows.Count + " firstEmpty=" + d.FirstEmpty + " nextId=" + d.NextId + " del=" + del + "\r\n");
                FbState st = FbState.Load();
                List<string> pk = new List<string>(); foreach (object k in st.Pend.Keys) pk.Add((string)k); pk.Sort(StringComparer.Ordinal);
                foreach (string k in pk) { PendItem pi = (PendItem)st.Pend[k]; o.Append("FP " + k + " t=" + pi.T + " h=" + pi.H + " n=" + (pi.N ? "True" : "False") + "\r\n"); }
                t = Tramite.ById("requerimiento");
                s = Store.Read(t);
                for (int i = 0; i < s.Clients.Count; i++) { ClientRec c = s.Clients[i]; c.F["xF"] = new DateTime(2026, 11, 1).AddDays(3 * i); c.F["xH"] = null; c.F["xI"] = i == 2 ? "Cerrado" : null; }
                d = Store.BuildData(t, s);
                for (int day = 1; day <= 30; day++) { int due, late; Store.ReqAlerts(d.Rows, new DateTime(2026, 11, day), out due, out late); o.Append("RQ " + day + " due=" + due + " late=" + late + "\r\n"); }
            }
            // etapa 3b: tasas DGT, IVA, comisión y margen de la ficha
            Config.Cur.TasasCanje = "B=33; C=100; D=100";
            foreach (string tid in new string[] { "canje", "traduccion", "homologacion" })
            {
                Tramite t = Tramite.ById(tid); StoreData s = Store.Read(t); int i = 0;
                foreach (ClientRec c in s.Clients)
                {
                    Dictionary<string, object> v = new Dictionary<string, object>(c.F);
                    if (i % 3 == 1) { v["comtipo"] = "%"; v["comval"] = 10L; } else if (i % 3 == 2) { v["comtipo"] = "EUR"; v["comval"] = "12,5"; }
                    i++;
                    double tx = Ps.TasasFor(t, v); double? p = Ps.Num(Store.Get(v, "price"));
                    StringBuilder line = new StringBuilder("MG " + tid + " " + c.Row + " tasas=" + N(tx));
                    if (p.HasValue && p.Value > 0) { double b, iv, t2; Ps.SplitIva(p.Value, tx, out b, out iv, out t2); line.Append(" base=" + N(b) + " iva=" + N(iv) + " tx=" + N(t2)); }
                    double m, pct, com;
                    if (Ps.Margin(t, v, out m, out pct, out com)) line.Append(" m=" + N(m) + " pct=" + pct.ToString("0.###", Ps.Inv) + " com=" + N(com)); else line.Append(" m=-");
                    o.Append(line.ToString() + "\r\n");
                }
            }
            // etapa 3b: la ficha (misma XAML): cada cliente se abre para editar; valores leídos de la ficha y textos de la tarjeta
            G.T = Tramite.ById("canje"); G.Cache = null;
            MainWin w = new MainWin();
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = Store.Read(t); if (s == null) continue;
                G.T = t; TramiteData d = Store.BuildData(t, s);
                G.Lists = d.Lists; G.Tariff = d.Tariff; G.DefaultPrice = d.DefaultPrice; G.Rows = d.Rows; G.NextId = d.NextId;
                w.BuildForm(); w.ResetForm();
                o.Append(FormLine("FN " + t.Id, t, w.GetFormValues()) + "\r\nFC " + t.Id + " new " + Card(w) + "\r\n");
                foreach (ClientRow r in d.Rows)
                {
                    w.FillFormForEdit(r);
                    o.Append(FormLine("FE " + t.Id + " " + r.Row, t, w.GetFormValues()) + "\r\nFC " + t.Id + " " + r.Row + " " + Card(w) + "\r\n");
                }
            }
            // etapa 3c: menú del clic derecho, atajos, responsable, eliminar y deshacer (hora fija, como pstest.ps1)
            App.Now = delegate { return new DateTime(2026, 11, 19, 12, 0, 0); };
            App.User = "PRUEBA"; App.Quiet = true;
            Config.Cur.Colleagues = new List<string>(new string[] { "ana", " LUIS", "ANA" }); Config.Cur.OpenAfterSave = false;
            Config.Cur.TabsChecked = new List<string>(); foreach (Tramite t in Tramite.All) Config.Cur.TabsChecked.Add(t.Id);   // Load-Data sin Excel
            w.ToastHook = delegate(string a, string b) { o.Append("TOAST " + a + "|" + b + "\r\n"); };
            w.ErrorHook = delegate(string m) { o.Append("ERR " + PsH.Rep(m, @"\s+", " ") + "\r\n"); };
            DataGrid grid = w.U<DataGrid>("ClientsGrid");
            foreach (Tramite t in Tramite.All)
            {
                if (Store.Read(t) == null) continue;
                G.T = t; Data.LoadData(false, false);
                foreach (ClientRow r in G.Rows) { grid.SelectedItem = r; w.FillCmMenu(r); o.Append("CM " + t.Id + " " + r.Row + " " + w.CmMenuText() + "\r\n"); }
            }
            Tramite tc = Tramite.ById("canje"); G.T = tc; Data.LoadData(false, false); w.BuildForm(); w.ResetForm();
            ClientRow q = TestRow(0); int qn = q.Row;
            List<string> sts = new List<string>(); foreach (string x in MainWin.GetList(tc.FieldByKey("status").List)) if (Ps.Str(x).Trim().Length > 0 && x != Ps.Str(q.Status)) sts.Add(x);
            grid.SelectedItem = q; w.QuickSetField("status", sts[sts.Count - 1], false); o.Append(RowText("QS", qn) + " undo=" + w.Undo.Key + "/" + Ps.Str(w.Undo.Old) + "\r\n");
            w.InvokeQuickUndo(); o.Append(RowText("QU", qn) + "\r\n");
            grid.SelectedItem = RowAt(qn); w.QuickSetField("paystate", "Pagado", false); o.Append(RowText("QS", qn) + "\r\n");
            w.InvokeQuickUndo(); o.Append(RowText("QU", qn) + "\r\n");
            w.InvokeQuickUndo();
            grid.SelectedItem = RowAt(qn); w.SetRespBulk("ANA", false); o.Append(RowText("RS", qn) + "\r\n");
            w.InvokeQuickUndo(); o.Append(RowText("RU", qn) + "\r\n");
            ClientRow q2 = TestRow(1); int q2n = q2.Row;
            w.RemoveClient(q2, false); o.Append(RowText("DL", q2n) + "\r\n");
            w.InvokeQuickUndo(); o.Append(RowText("DU", q2n) + "\r\n");
            // pestaña hecha a mano (tasas)
            Tramite tt = Tramite.ById("tasas"); G.T = tt; Data.LoadData(false, false); w.BuildForm(); w.ResetForm();
            grid.SelectedItem = HandRows()[0]; w.SetRespBulk("LUIS", false); string nid = w.Undo.Items[0].Id;
            ClientRow nr = null; foreach (ClientRow r in G.Rows) if (r.Id == nid) { nr = r; break; }
            int nrow = nr.Row; o.Append(RowText("RH", nrow) + "\r\n");
            w.InvokeQuickUndo(); o.Append(RowText("RHU", nrow) + "\r\n");
            foreach (int i in new int[] { 1, 2 })
            {
                w.RemoveClient(HandRows()[i], false);
                o.Append("DH " + i + " n=" + G.Rows.Count + "\r\n");
                if (i == 1) { w.InvokeQuickUndo(); o.Append("DHU n=" + G.Rows.Count + "\r\n"); }
            }
            App.Now = delegate { return DateTime.Now; };
            // etapa 3d: documentos (carpeta de prueba de docsfixture.ps1)
            Config.Cur.DocsDir = Path.Combine(testDir, "cs_docs"); string droot = Config.Cur.DocsDir;
            DocsTree(o, "DT0", droot);
            Docs.RepairAll();
            DocsTree(o, "DT1", droot);
            G.T = Tramite.ById("canje"); Data.LoadData(false, false);
            foreach (ClientRow r in G.Rows) { int n = Docs.Count(r); if (n > 0) o.Append("DM " + r.Row + " " + r.Id + " n=" + n + "\r\n"); }
            string[] ids = File.ReadAllText(Path.Combine(droot, "clientes.txt"), Encoding.UTF8).Trim().Split('|');
            foreach (int k in new int[] { 0, 1, 4, 7, 8 }) { w.OpenDocsPanel(RowById(ids[k])); o.Append("DP " + k + " " + DocsPanelText(w) + "\r\n"); }
            w.OpenDocsPanel(RowById(ids[0])); w.AddClientDocs(new string[] { Path.Combine(testDir, "docs_src\\nuevo.pdf"), Path.Combine(testDir, "docs_src\\a.pdf"), Path.Combine(testDir, "docs_src\\no-existe.pdf") });
            o.Append("DA 0 " + DocsPanelText(w) + "\r\n");
            w.OpenDocsPanel(RowById(ids[8])); w.AddClientDocs(new string[] { Path.Combine(testDir, "docs_src\\nuevo.pdf") });
            o.Append("DA 8 " + DocsPanelText(w) + "\r\n");
            foreach (ClientRow r in G.Rows) { int n = Docs.Count(r); if (n > 0) o.Append("DM2 " + r.Row + " " + r.Id + " n=" + n + "\r\n"); }
            DocsTree(o, "DT2", droot);
            // etapa 4a: Resumen (todas las pestañas en caché; hora fija; sin animación de cifras)
            App.Now = delegate { return new DateTime(2026, 11, 19, 12, 0, 0); };
            G.Cache = new Dictionary<string, TramiteData>();
            foreach (Tramite t in Tramite.All) { StoreData s = Store.Read(t); if (s != null) G.Cache[t.Id] = Store.BuildData(t, s); }
            foreach (string sc in new string[] { null, "canje", "tasas", "traduccion" })
            {
                w.StatsScopeId = sc; w.FillStatsFilters();
                o.Append("SF " + sc + " who=" + string.Join(",", Items(w.U<ComboBox>("StatsWho"))) + "|scope=" + w.U<ComboBox>("StatsScope").SelectedIndex + "/" + w.U<ComboBox>("StatsScope").Items.Count + "|per=" + string.Join(",", Items(w.U<ComboBox>("StatsPeriod"))) + "\r\n");
                foreach (int p in new int[] { 0, 1, 2, 3 }) foreach (int wi in new int[] { 0, 1, 3 })
                {
                    w.U<ComboBox>("StatsPeriod").SelectedIndex = p; w.U<ComboBox>("StatsWho").SelectedIndex = wi; w.UpdateStats();
                    o.Append("ST " + sc + " p=" + p + " w=" + wi + " " + StatsText(w) + "\r\n");
                }
            }
            w.StatsScopeId = null; w.FillStatsFilters(); w.U<ComboBox>("StatsPeriod").SelectedIndex = 0; w.U<ComboBox>("StatsWho").SelectedIndex = 0;
            foreach (string m in new string[] { "ChartDay", "ChartWeek", "ChartMonth" }) { w.U<RadioButton>(m).IsChecked = true; w.UpdateChart(); o.Append("CH " + m + " " + StatsText(w) + "\r\n"); }
            App.Now = delegate { return DateTime.Now; };
            // etapa 4b: Ajustes (mismos valores de partida que en pstest.ps1; la copia externa no cambia: Backup-External es de la etapa 5)
            App.DataDir = Path.Combine(testDir, "cs_data"); Directory.CreateDirectory(App.DataDir);
            Config.Cur.CompanyName = ""; Config.Cur.CompanyLogo = ""; Config.Cur.InvIban = ""; Config.Cur.ExtBackupDir = ""; Config.Cur.LastExtBackup = "";
            Config.Cur.InvName = "Ana Prueba"; Config.Cur.InvActivity = "Gestoría"; Config.Cur.InvNif = "x1"; Config.Cur.InvAddress = "C/ Mayor 1"; Config.Cur.InvTown = "Madrid";
            Config.Cur.TasasCanje = "B=33; C=100; D=100"; Config.Cur.Backup = true; Config.Cur.OpenAfterSave = false; Config.Cur.DefaultNationality = "MARROQUI"; Config.Cur.Colleagues = new List<string>(new string[] { "SAFAE", "CHAYMAE" });
            w.LoadSettingsPanel(); o.Append("SP0 " + SetText(w) + "\r\n");
            w.U<TextBox>("CompanyBox").Text = "  Gestoría Prueba  "; w.U<TextBox>("InvNifBox").Text = " x123 "; w.U<TextBox>("InvIbanBox").Text = "es12 3456"; w.U<TextBox>("InvPobBox").Text = " Getafe ";
            w.U<TextBox>("TasasBox").Text = "B=40;C=90\r\n  D=95 ;; "; w.U<TextBox>("InvNextBox").Text = " 7 "; w.U<TextBox>("ColleaguesBox").Text = "ana, luis ,, Pepa"; w.U<ComboBox>("NatBox").Text = " argelina ";
            w.U<CheckBox>("BackupChk").IsChecked = false; w.U<CheckBox>("OpenAfterChk").IsChecked = true;
            string lsrc = Path.Combine(testDir, "logo_cs.png"); File.WriteAllBytes(lsrc, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")); w.PendingLogo = lsrc;
            w.U<TextBox>("DocsDirBox").Text = Path.Combine(testDir, "cs_docs2");
            w.SaveSettings();
            o.Append("SC1 " + CfgText(w) + "\r\n"); o.Append("SP1 " + SetText(w) + "\r\n");
            DocsTree(o, "SD", Path.Combine(testDir, "cs_docs2"));
            w.U<TextBox>("InvNextBox").Text = "abc"; w.U<TextBox>("CompanyBox").Text = ""; w.PendingLogo = ""; w.SaveSettings();
            o.Append("SC2 " + CfgText(w) + "\r\n"); o.Append("SP2 " + SetText(w) + "\r\n");            // etapa 4c: sesión
            Config.Cur.Colleagues = new List<string>(new string[] { "SAFAE", "chaymae", " ", "LUIS" }); Config.Cur.LastUser = "luis"; G.Lists["colleagues"] = new string[] { "SAFAE", "", "X" };
            // al iniciar sesión se escribe «actividad GESTEM - PC.txt» en la carpeta externa (registro propio de la prueba; fecha y hora fuera)
            string ssExt = Path.Combine(testDir, "cs_sess_ext"), ssKeepExt = Config.Cur.ExtBackupDir, ssKeepLog = App.LogPath; Directory.CreateDirectory(ssExt);
            App.LogPath = Path.Combine(testDir, "sess_log_cs.txt"); if (File.Exists(App.LogPath)) File.Delete(App.LogPath); Config.Cur.ExtBackupDir = ssExt; bool ssOwner = App.IsOwnerPC, ssAdm = App.IsAdmin; App.IsOwnerPC = false; App.IsAdmin = false;    // como pstest: aún no es el PC del creador
            Func<string> ssAct = delegate
            {
                string f = Path.Combine(ssExt, "actividad GESTEM - " + Environment.MachineName + ".txt"); if (!File.Exists(f)) return "no";
                string[] a = File.ReadAllLines(f, Encoding.UTF8); List<string> l = new List<string>();
                for (int i = 0; i < a.Length; i++) l.Add(i == 1 ? "Updated" : (i >= 3 && a[i].Length > 16 ? a[i].Substring(16) : a[i]));
                File.Delete(f); return string.Join(" ¶ ", l.ToArray());
            };
            w.U<ContentControl>("UserBtn").Content = ""; w.ShowSession(); o.Append("SS0 " + SessText(w) + "\r\n");
            foreach (string who in new string[] { " pepa ", "", "Safae", "luis" }) { w.StartSession(who); o.Append("SS1 [" + who + "] " + SessText(w) + "\r\n"); o.Append("SSA [" + who + "] " + ssAct() + "\r\n"); }
            w.ShowSession(); o.Append("SS2 " + SessText(w) + "\r\n");
            Config.Cur.LastUser = ""; App.User = ""; w.StartSession("  "); o.Append("SS3 " + SessText(w) + "\r\n"); o.Append("SSA3 " + ssAct() + "\r\n");
            w.ShowSession(); o.Append("SS4 " + SessText(w) + "\r\n"); Config.Cur.ExtBackupDir = ssKeepExt; App.LogPath = ssKeepLog; App.IsOwnerPC = ssOwner; App.IsAdmin = ssAdm;
            // etapa 4d: administrador (registro de prueba fijo; admin.dat en la carpeta de datos de prueba)
            MainWin.Admin = new AdminData(); App.IsAdmin = false; App.IsOwnerPC = MainWin.TestOwnerPC(); w.Remote = null; w.AdmLogFilter = "all";
            bool keepFb = Config.Cur.FbSync; Config.Cur.FbSync = false;
            o.Append("AH " + MainWin.PwHash("secreto1", "sal") + " " + App.IsOwnerPC + "\r\n");
            string keepLog = App.LogPath; App.LogPath = Path.Combine(testDir, "adm_log_cs.txt"); File.WriteAllBytes(App.LogPath, Convert.FromBase64String("MjAyNi0xMS0xOCAwOTowMDowMSAgW1NBRkFFXSBzZXNpw7NuIGluaWNpYWRhDQoyMDI2LTExLTE4IDA5OjAxOjAwICBbU0FGQUVdIGHDsWFkacOzIGNsaWVudGUgQ0FOSi0xIMK3IEFOQSBSVUlaIMK3IEIgKENvY2hlKQ0KMjAyNi0xMS0xOCAwOTowMjowMCAgW0FETUlOXSBhZG1pbjogYWNjZXNvDQoyMDI2LTExLTE4IDA5OjAzOjAwICBbXSBjb3BpYSBleHRlcm5hOiBEOlx4DQoyMDI2LTExLTE4IDA5OjA0OjAwICBbTFVJU10gTG9hZC1EYXRhOiBhbGdvDQoyMDI2LTExLTE4IDA5OjA1OjAwICBbXSBVSSBFUlJPUiAodmVudGFuYSkgTnVsbFJlZmVyZW5jZQ0KMjAyNi0xMS0xOCAwOTowNjowMCAgRkFUQUwgIHNlIGNlcnLDsw0KMjAyNi0xMS0xOCAwOTowNzowMCAgW2FkbWluXSBhZG1pbjogaW50ZW50byBkZSBhY2Nlc28gZmFsbGlkbw0KbMOtbmVhIHJvdGENCjIwMjYtMTEtMTggMDk6MDg6MDAgIFtDSEFZTUFFXSBlZGl0w7MgY2xpZW50ZSBYIMK3IFkgwrcgWg0KMjAyNi0xMS0xOCAwOTowOTowMCAgW10gQ29udHJhc2XDsWEgaW5jb3JyZWN0YS4gb3RyYSB2ZXoNCjIwMjYtMTEtMTggMDk6MTA6MDAgIFtdIERFQlVHIGVycm9yOiB4DQoyMDI2LTExLTE4IDA5OjExOjAwICBbUEVQQV0gbG9nbzogbm8gc2UgcHVkbw0K"));
            foreach (string lang in new string[] { "es", "fr" }) { Config.Cur.Language = lang; Tx.Lang = lang; foreach (string fl in new string[] { "all", "team", "admin", "errors" })
            {
                w.AdmLogFilter = fl; List<AdmLogItem> its = w.FillAdminLog(); List<string> l = new List<string>();
                foreach (AdmLogItem it in its) l.Add(it.Time + "|" + it.User + "|" + it.UserBg + "|" + it.UserFg + "|" + it.Text + "|" + it.Fg);
                o.Append("AL " + lang + " " + fl + " title=" + w.U<TextBlock>("AdmLogTitle").Text + "|err=" + Ps.Str(w.U<ContentControl>("AdmLogErr").Content) + "|" + string.Join(" ; ", l) + "\r\n");
            } }
            App.LogPath = keepLog; w.AdmLogFilter = "all";
            foreach (string lang in new string[] { "fr", "es" }) { Config.Cur.Language = lang; Tx.Lang = lang; w.ApplyAdminLanguage(); List<string> l = new List<string>(); AdmTree(w.U<FrameworkElement>("PanelAdmin"), l); AdmTree(w.U<FrameworkElement>("PanelAdminLogin"), l); o.Append("AT " + lang + " " + string.Join(" | ", l) + "\r\n"); }
            Func<string, PasswordBox> PB = delegate(string n) { return w.U<PasswordBox>(n); };
            w.ShowAdminLogin(); o.Append("AD0 " + AdmState(w) + "\r\n");
            foreach (string[] pp in new string[][] { new string[] { "abc", "" }, new string[] { "abcdef", "abcdeX" }, new string[] { "abcdef", "abcdef" } }) { PB("AdmPw").Password = pp[0]; PB("AdmPw2").Password = pp[1]; w.InvokeAdminLogin(); o.Append("AD1 " + AdmState(w) + "\r\n"); }
            w.CloseAdmin(); w.ShowAdminLogin(); o.Append("AD2 " + AdmState(w) + "\r\n");
            PB("AdmPw").Password = "mala12"; w.InvokeAdminLogin(); o.Append("AD3 " + AdmState(w) + "\r\n");
            PB("AdmPw").Password = "abcdef"; w.InvokeAdminLogin(); o.Append("AD4 " + AdmState(w) + "\r\n");
            w.U<CheckBox>("AdmLockFile").IsChecked = true; w.U<CheckBox>("AdmLockAll").IsChecked = true; w.U<CheckBox>("AdmLockTeam").IsChecked = true; w.U<CheckBox>("AdmIgnorePause").IsChecked = true; w.SaveAdminLocks(); o.Append("AD5 " + AdmState(w) + "\r\n");
            w.CloseAdmin(); o.Append("AD6 " + AdmState(w) + "\r\n");
            MainWin.Admin = new AdminData(); MainWin.LoadAdmin();
            List<string> lk = new List<string>(); foreach (object k in MainWin.Admin.Locks.Keys) lk.Add((string)k); lk.Sort(StringComparer.CurrentCultureIgnoreCase);
            for (int i = 0; i < lk.Count; i++) lk[i] = lk[i] + "=" + MainWin.Admin.Locks[lk[i]];
            o.Append("AD7 locks=" + string.Join(",", lk) + "|ign=" + MainWin.Admin.IgnorePause + "|ok=" + (MainWin.PwHash("abcdef", MainWin.Admin.Salt) == MainWin.Admin.Hash) + "\r\n");
            App.IsAdmin = true; w.OpenAdminPanel(); o.Append("AD8 " + AdmState(w) + "|info=" + System.Text.RegularExpressions.Regex.Replace(w.U<TextBlock>("AdmInfo").Text, "·.*$", "") + "|remote=" + w.U<TextBlock>("AdmRemoteTxt").Text + "\r\n");
            foreach (string[] pp in new string[][] { new string[] { "12", "12" }, new string[] { "nueva123", "nueva124" }, new string[] { "nueva123", "nueva123" } }) { PB("AdmNewPw").Password = pp[0]; PB("AdmNewPw2").Password = pp[1]; w.ChangeAdminPw(); o.Append("AD9 " + AdmState(w) + "\r\n"); }
            w.CloseAdmin(); MainWin.Admin.IgnorePause = false;
            w.Remote = Json.Parse("{\"mensaje\":\"hola\",\"pausa\":true,\"motivo\":\"mantenimiento\"}"); w.ApplyPause(); w.UpdateRemoteText();
            o.Append("AR paused=" + w.U<FrameworkElement>("PanelPaused").Visibility + "|" + w.U<TextBlock>("PausedText").Text + "|" + w.U<TextBlock>("AdmRemoteTxt").Text + "\r\n");
            w.Remote = Json.Parse("{\"pausa\":\"false\"}"); w.ApplyPause(); w.UpdateRemoteText(); o.Append("AR2 paused=" + w.U<FrameworkElement>("PanelPaused").Visibility + "|" + w.U<TextBlock>("AdmRemoteTxt").Text + "\r\n");
            // un PC que no es el del creador: los bloqueos esconden partes de Ajustes y la pausa remota se ve
            App.IsOwnerPC = false; MainWin.Admin.OwnerPC = false; App.IsAdmin = false; w.ApplyLocks(); o.Append("AD10 " + AdmState(w) + "\r\n");
            w.Remote = Json.Parse("{\"mensaje\":\"hola\",\"pausa\":true,\"motivo\":\"mantenimiento\"}"); w.ApplyPause(); o.Append("AR3 paused=" + w.U<FrameworkElement>("PanelPaused").Visibility + "|" + w.U<TextBlock>("PausedText").Text + "\r\n");
            w.Remote = Json.Parse("{\"pausa\":1}"); w.ApplyPause(); o.Append("AR4 paused=" + w.U<FrameworkElement>("PanelPaused").Visibility + "|" + w.U<TextBlock>("PausedText").Text + "\r\n");
            MainWin.Admin.Locks = PsH.New(); w.ApplyLocks();
            w.Remote = null; w.ApplyPause(); App.IsAdmin = false; Config.Cur.Language = "es"; Tx.Lang = "es"; Config.Cur.FbSync = keepFb;            // etapa 4e: Colaboradores (unos clientes de canje y autónomo con colaborador y comisión)
            string[][] fx = { new string[] { "canje", "0", "abogado x", "%", "10", "" }, new string[] { "canje", "1", "ABOGADO X", "", "25", "Pagada" }, new string[] { "canje", "2", "Luis ", "%", "12,5", "" }, new string[] { "canje", "6", "luis", "", "", "" }, new string[] { "autonomo", "0", "abogado x", "", "30", "" }, new string[] { "autonomo", "1", "MARTA", "%", "x", "" } };
            foreach (string tid in new string[] { "canje", "autonomo" })
            {
                Tramite t = Tramite.ById(tid); StoreData s = Store.Read(t);
                foreach (string[] f in fx) if (f[0] == tid) { ClientRec cl = s.Clients[int.Parse(f[1])]; cl.F["colab"] = f[2]; cl.F["comtipo"] = f[3].Length > 0 ? f[3] : null; cl.F["comval"] = f[4].Length > 0 ? f[4] : null; cl.F["compag"] = f[5].Length > 0 ? f[5] : null; }
                Store.Write(t, s);
            }
            App.Now = delegate { return new DateTime(2026, 11, 19, 12, 0, 0); };
            List<string> keepCol = Config.Cur.Collaborators; Config.Cur.Collaborators = new List<string> { "ZETA", "abogado x" };
            G.T = Tramite.ById("canje"); Data.LoadData(true, true);
            w.ApplyLanguage();
            List<string> hd = new List<string>(); foreach (DataGridColumn col in w.U<DataGrid>("ColabGrid").Columns) hd.Add(Ps.Str(col.Header));
            o.Append("CO texts " + w.U<TextBlock>("ColabTitle").Text + "|" + w.U<TextBlock>("ColabHint").Text + "|" + w.U<ContentControl>("ColabPaidBtn").Content + "|" + w.U<ContentControl>("ColabOnlyPend").Content + "|" + w.U<TextBlock>("ColK1L").Text + "|" + w.U<TextBlock>("ColK6L").Text + "|" + string.Join(",", hd) + "\r\n");
            w.LoadColabPanel(); o.Append("CO0 " + ColState(w) + "\r\n");
            w.U<ListBox>("ColabList").SelectedIndex = 1; w.UpdateColabView(); o.Append("CO1 " + ColState(w) + "\r\n");
            w.U<CheckBox>("ColabOnlyPend").IsChecked = true; w.UpdateColabView(); o.Append("CO2 " + ColState(w) + "\r\n");
            w.LoadColabPanel(); o.Append("CO3 " + ColState(w) + "\r\n");
            w.U<CheckBox>("ColabOnlyPend").IsChecked = false; w.U<ListBox>("ColabList").SelectedIndex = 0; w.UpdateColabView(); w.U<DataGrid>("ColabGrid").SelectAll(); w.MarkColabPaid(); o.Append("CO4 " + ColState(w) + "\r\n");
            w.InvokeQuickUndo(); w.LoadColabPanel(); o.Append("CO5 " + ColState(w) + "\r\n");
            Config.Cur.Collaborators = new List<string>(); G.Cache = new Dictionary<string, TramiteData>(StringComparer.CurrentCultureIgnoreCase); w.LoadColabPanel(); o.Append("CO6 " + ColState(w) + "\r\n");
            Config.Cur.Collaborators = keepCol; G.T = Tramite.ById("canje"); Data.LoadData(true, true); App.Now = delegate { return DateTime.Now; };            // etapa 5a: copia externa (carpeta de prueba; 62 copias viejas para ver que quedan 60) y registro de actividad sin nombres
            { string company = ExtBackup.FindCompanyCloud(); o.Append("EX cloud=" + string.Join(";", ExtBackup.CloudRoots()) + "|company=" + company + "\r\n"); }
            App.Now = delegate { return new DateTime(2026, 11, 19, 12, 0, 0); };
            string ext = Path.Combine(testDir, "cs_ext"); if (Directory.Exists(ext)) Directory.Delete(ext, true); Directory.CreateDirectory(ext);
            for (int i = 1; i <= 62; i++) File.WriteAllText(Path.Combine(ext, "x_" + new DateTime(2026, 1, 1).AddDays(i).ToString("yyyy-MM-dd") + ".xlsx"), "viejo");
            File.WriteAllText(Path.Combine(ext, "x_2026-1-1.xlsx"), "otro");
            string keepLog2 = App.LogPath; App.LogPath = Path.Combine(testDir, "ext_log_cs.txt"); File.WriteAllBytes(App.LogPath, Convert.FromBase64String("MjAyNi0xMS0xOCAwOTowMDowMSAgW1NBRkFFXSBzZXNpw7NuIGluaWNpYWRhDQoyMDI2LTExLTE4IDA5OjAxOjAwICBbU0FGQUVdIGHDsWFkacOzIGNsaWVudGUgQ0FOSi0yMDI2LTAwMSDCtyBBTkEgUlVJWiDCtyBCIChDb2NoZSkNCjIwMjYtMTEtMTggMDk6MDE6MzAgIFtsdWlzXSBlZGl0w7MgY2xpZW50ZSBBVVQtMSDCtyBQRVBFIMK3IFggwrcgQXV0w7Nub21vDQoyMDI2LTExLTE4IDA5OjAyOjAwICBbYWRtaW5dIGFkbWluOiBhY2Nlc28NCjIwMjYtMTEtMTggMDk6MDM6MDAgIFtdIGNvcGlhIGV4dGVybmE6IEQ6XHgNCjIwMjYtMTEtMTggMDk6MDQ6MDAgIFtMVUlTXSBMb2FkLURhdGE6IGFsZ28NCjIwMjYtMTEtMTggMDk6MDQ6MzAgIFtMVUlTXSBERUJVRyBlcnJvcjogeA0KMjAyNi0xMS0xOCAwOTowNTowMCAgW10gVUkgRVJST1IgKHZlbnRhbmEpIE51bGxSZWZlcmVuY2UNCmJhc3VyYQ0KMjAyNi0xMS0xOCAwOTowNjowMCAgW0NIQVlNQUVdIGNhbWJpw7MgbG9zIGFqdXN0ZXMNCjIwMjYtMTEtMTggMDk6MDc6MDAgIFtDSEFZTUFFXSBjb21pc2nDs24gcGFnYWRhOiAyIHRyw6FtaXRlKHMpIMK3IExVSVMNCg=="));
            Config.Cur.ExtBackupDir = ext; Config.Cur.LastExtBackup = "";
            bool extOk = w.BackupExternal(true);
            List<string> efl = new List<string>(); foreach (string f in Directory.GetFiles(ext, "*", SearchOption.AllDirectories)) efl.Add(System.Text.RegularExpressions.Regex.Replace(f.Substring(ext.Length + 1), @"^documentos\\cs_", @"documentos\X_"));
            Ps.Sort(efl, StringComparer.CurrentCultureIgnoreCase.Compare);
            int nx = 0; string firstX = null; foreach (string f in efl) if (System.Text.RegularExpressions.Regex.IsMatch(f, @"^x_", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) { if (firstX == null) firstX = f; if (System.Text.RegularExpressions.Regex.IsMatch(f, @"^x_.*\.xlsx$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) nx++; }
            o.Append("EX ok=" + extOk + " last=" + Config.Cur.LastExtBackup + " n=" + efl.Count + " excel=" + nx + " first=" + firstX + "\r\n");
            foreach (string f in efl) if (!System.Text.RegularExpressions.Regex.IsMatch(f, "^x_", System.Text.RegularExpressions.RegexOptions.IgnoreCase)) o.Append("EXF " + f + "\r\n");
            o.Append("EXA " + string.Join(" ¶ ", File.ReadAllLines(Path.Combine(ext, "actividad GESTEM - " + Environment.MachineName + ".txt"), Encoding.UTF8)) + "\r\n");
            Config.Cur.ExtBackupDir = Path.Combine(testDir, "cs_x.xlsx"); o.Append("EX2 ok=" + w.BackupExternal(true) + "\r\n");
            Config.Cur.ExtBackupDir = ""; Config.Cur.LastExtBackup = ""; App.LogPath = keepLog2; App.Now = delegate { return DateTime.Now; };
            // etapa 5b: control.json y versión con un GitHub de mentira
            object[] fake = null;
            Updater.FetchHook = delegate(string name) { object r = fake[0]; if (r is int) { if ((int)r == 0) throw new Exception("sin conexión"); throw new RepoException((int)r, "(" + r + ")"); } return (string)r; };
            w.InfoHook = delegate(string m) { o.Append("INFO " + m + "\r\n"); };
            foreach (string f in new string[] { "control.cache", "control.visto" }) File.Delete(Path.Combine(App.DataDir, f));
            w.Remote = null; App.IsAdmin = false; App.IsOwnerPC = false; MainWin.Admin.IgnorePause = false;
            int rn = 0; foreach (object r in new object[] { "{\"mensaje\":\" Hola equipo \",\"pausa\":true,\"motivo\":\"obras\"}", "{\"mensaje\":\"Hola equipo\",\"pausa\":true,\"motivo\":\"obras\"}", 0, "{oops", 500, 404, "{\"pausa\":false,\"mensaje\":\"otro\"}", 0 })
            { fake = new object[] { r }; w.CheckRemote(); o.Append("RC" + rn + " " + RemState(w) + "\r\n"); rn++; }
            foreach (string v in new string[] { "1.9.3", "1.9.2", "1.10", "2.0.0.1", "1.9.3.1", "1.9.3.0", "hola", "1", "1.2.3.4.5", " 1.9.4 " }) o.Append("RV " + v + "=" + Updater.IsNewer(v, App.Version) + "\r\n");
            foreach (object v in new object[] { "1.9.3", "hola", 0, 401 }) { fake = new object[] { v }; w.CheckUpdate(true); o.Append("RU " + v + "\r\n"); }
            { string keepKey = Config.Cur.UpdateKey; Updater.SetKey("clave-xyz"); o.Append("RK " + Updater.GetKey() + "|" + (Config.Cur.UpdateKey.Length > 0)); Updater.SetKey(""); o.Append("|" + Updater.GetKey() + "|" + Config.Cur.UpdateKey + "\r\n"); Config.Cur.UpdateKey = keepKey; }
            Updater.FetchHook = null; w.InfoHook = null; w.Remote = null; w.ApplyPause();
            // etapa 5c: carpeta de documentos en la nube (dos «nubes» de prueba: la más reciente gana; más hondo de 3 o oculta no cuenta)
            {
                string[] nb = { Path.Combine(testDir, "cs_nube1"), Path.Combine(testDir, "cs_nube2") };
                App.CloudRootsOverride = new List<string>(nb);
                foreach (string d in nb) if (Directory.Exists(d)) { foreach (string x in Directory.GetDirectories(d, "*", SearchOption.AllDirectories)) File.SetAttributes(x, FileAttributes.Directory); Directory.Delete(d, true); }
                Func<string, DateTime, DirectoryInfo> mk = delegate(string p, DateTime t) { DirectoryInfo x = Directory.CreateDirectory(p); x.LastWriteTime = t; return x; };
                mk(Path.Combine(nb[0], @"a\GESTEM documentos"), new DateTime(2026, 5, 1));
                mk(Path.Combine(nb[1], @"x\y\z\GESTEM documentos"), new DateTime(2026, 6, 1));
                mk(Path.Combine(nb[1], @"a\b\c\d\GESTEM documentos"), new DateTime(2026, 9, 1));
                DirectoryInfo hidDir = mk(Path.Combine(nb[0], "oculta"), DateTime.Now); hidDir.Attributes = FileAttributes.Hidden | FileAttributes.Directory; mk(Path.Combine(hidDir.FullName, "GESTEM documentos"), new DateTime(2026, 10, 1));
                Func<string, string> mask = delegate(string p) { return p != null ? System.Text.RegularExpressions.Regex.Replace(p.Substring(testDir.Length + 1), "^cs_", "X_") : "-"; };
                o.Append("DC find=" + mask(ExtBackup.FindCloudDocsDir()) + "\r\n");
                string keepDocs = Config.Cur.DocsDir; bool keepFb5c = Config.Cur.FbSync;
                string loc = Docs.LocalRoot(); if (Directory.Exists(loc)) Directory.Delete(loc, true); Directory.CreateDirectory(Path.Combine(loc, @"DGT ACTUALIZADO\CANJ-1"));
                File.WriteAllText(Path.Combine(loc, @"DGT ACTUALIZADO\CANJ-1\a.pdf"), "uno"); File.WriteAllText(Path.Combine(loc, "b.pdf"), "dos");
                Action<string, string> keepToast = w.ToastHook; w.ToastHook = delegate(string a, string b) { o.Append("TOAST " + a.Replace(testDir + "\\cs_", "X_") + "|" + b + "\r\n"); };
                Config.Cur.DocsDir = ""; Config.Cur.FbSync = false; w.InitializeDocsFolder(); o.Append("DC off docs=" + Config.Cur.DocsDir + "\r\n");
                Config.Cur.FbSync = true; w.InitializeDocsFolder();
                List<string> dl = new List<string>(); foreach (string f in Directory.GetFiles(Config.Cur.DocsDir, "*", SearchOption.AllDirectories)) dl.Add(f.Substring(Config.Cur.DocsDir.Length + 1)); Ps.Sort(dl, StringComparer.CurrentCultureIgnoreCase.Compare);
                o.Append("DC on docs=" + mask(Config.Cur.DocsDir) + " local=" + Directory.Exists(loc) + " files=" + string.Join(",", dl) + "\r\n");
                foreach (string d in nb) { File.SetAttributes(Path.Combine(nb[0], "oculta"), FileAttributes.Directory); }
                foreach (string d in nb) if (Directory.Exists(d)) Directory.Delete(d, true);
                Config.Cur.DocsDir = ""; w.InitializeDocsFolder(); o.Append("DC none docs=" + Config.Cur.DocsDir + "\r\n");
                Config.Cur.DocsDir = keepDocs; Config.Cur.FbSync = keepFb5c; App.CloudRootsOverride = null; w.ToastHook = keepToast; Docs.UpdateMap();
            }
            // etapa 5d: Firebase sin red: las mismas vueltas (unión, envío, recibir) con resultados de mentira
            {
                Action<string, string> keepToast5 = w.ToastHook; w.ToastHook = delegate(string a, string b) { o.Append("TOAST " + System.Text.RegularExpressions.Regex.Replace(a, "[0-9a-f]{32}", "GUID") + "|" + b + "\r\n"); };
                Fb.Notify = delegate(string m) { w.ShowToast(m, "", null, 15); }; Fb.LastErr = ""; Fb.Auth = null;
                if (File.Exists(FbState.Path)) File.Delete(FbState.Path);
                Func<string, List<FbDoc>> docs = delegate(string b64)
                {
                    List<FbDoc> l = new List<FbDoc>();
                    foreach (object x in (System.Collections.IList)Json.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(b64))))
                    { object del = Json.Get(x, "Deleted"); l.Add(new FbDoc { Uid = Json.Str(x, "Uid"), T = Json.Str(x, "T"), Id = Json.Str(x, "Id") ?? "", Datos = Json.Str(x, "Datos"), Updated = Json.Str(x, "Updated"), By = Json.Str(x, "By"), Pc = Json.Str(x, "Pc"), Deleted = del is bool && (bool)del }); }
                    return l;
                };
                List<FbPush> push = new List<FbPush>(); bool joinNext = true;
                Action<List<FbDoc>, string, bool> round = delegate(List<FbDoc> pulled, string mod, bool echo)
                {
                    // como Step-FbSync: lo que vuelve de la vuelta anterior, y lo que se enviaría en la siguiente
                    FbResult r = new FbResult(); r.Pulled = pulled; r.Mod = mod; if (echo) r.Pushed = push;
                    Fb.ApplyResult(r, joinNext);
                    FbState st = FbState.Load(); joinNext = st.Joined != "push" && st.Joined != "yes";
                    bool drop = false; push = joinNext ? new List<FbPush>() : Fb.BuildPush(st, out drop); if (!joinNext && drop) st.Save();
                };
                round(docs("W3siUGMiOiJPVFJPUEMiLCJEZWxldGVkIjpmYWxzZSwiVWlkIjoidS0wMDEiLCJJZCI6IkNBTkotMjAyNi0wMDEiLCJCeSI6Ik9UUkEiLCJUIjoiY2FuamUiLCJEYXRvcyI6IntcIm5hbWVcIjpcIk1PSEFNTUFESU5FIEJPVVRFUkZBU1NcIixcImRvY1wiOlwiWjQ4MDEwMDNCXCIsXCJzdGF0dXNcIjpcIkZpbmFsaXphZG9cIixcInByaWNlXCI6MTIwLFwicGF5c3RhdGVcIjpcIlBhZ2Fkb1wiLFwicGFpZFwiOjEyMH0iLCJVcGRhdGVkIjoiMjAyNi0xMS0yMFQxMDowMDowMCJ9LHsiUGMiOiJPVFJPUEMiLCJEZWxldGVkIjpmYWxzZSwiVWlkIjoidS0wMDIiLCJJZCI6IkNBTkotMjAyNi0wMDIiLCJCeSI6Ik9UUkEiLCJUIjoiY2FuamUiLCJEYXRvcyI6IntcIm5hbWVcIjpcIk9UUkEgUEVSU09OQVwiLFwiZG9jXCI6XCJYMTIzNDU2N1pcIixcInByaWNlXCI6MTQwfSIsIlVwZGF0ZWQiOiIyMDI2LTExLTIwVDEwOjAxOjAwIn0seyJQYyI6Ik9UUk9QQyIsIkRlbGV0ZWQiOmZhbHNlLCJVaWQiOiJ1LTA5MCIsIklkIjoiQ0FOSi0yMDI2LTA5MCIsIkJ5IjoiT1RSQSIsIlQiOiJjYW5qZSIsIkRhdG9zIjoie1wibmFtZVwiOlwiTlVFVkEgREUgRlVFUkFcIixcInByaWNlXCI6MTIuNSxcImZzb2xcIjpcIjIwMjYtMTEtMDJcIn0iLCJVcGRhdGVkIjoiMjAyNi0xMS0yMFQxMDowMjowMCJ9LHsiUGMiOiJPVFJPUEMiLCJEZWxldGVkIjp0cnVlLCJVaWQiOiJoYW5kLWNhbmplLWFiYyIsIklkIjoiIiwiQnkiOiJPVFJBIiwiVCI6ImNhbmplIiwiRGF0b3MiOiJ7XCJuYW1lXCI6XCJBTEdVSUVOIEEgTUFOT1wifSIsIlVwZGF0ZWQiOiIyMDI2LTExLTIwVDEwOjAzOjAwIn0seyJQYyI6Ik9UUk9QQyIsIkRlbGV0ZWQiOmZhbHNlLCJVaWQiOiJ1LWEwMSIsIklkIjoiQVVULTIwMjYtMDAxIiwiQnkiOiJPVFJBIiwiVCI6ImF1dG9ub21vIiwiRGF0b3MiOiJ7XCJuYW1lXCI6XCJNT0hBTU1FRCBEQUhCSVwiLFwieElcIjpcIk5PLURFQkUtTExFR0FSXCIsXCJzdGF0dXNcIjpcIkZhdm9yYWJsZVwifSIsIlVwZGF0ZWQiOiIyMDI2LTExLTIxVDA5OjAwOjAwIn0seyJQYyI6Ik9UUk9QQyIsIkRlbGV0ZWQiOmZhbHNlLCJVaWQiOiJ1LXp6IiwiSWQiOiJYIiwiQnkiOiJPVFJBIiwiVCI6ImluZXhpc3RlbnRlIiwiRGF0b3MiOiJ7fSIsIlVwZGF0ZWQiOiIyMDI2LTExLTIwVDEwOjAwOjAwIn1d"), "2026-11-20T09:00:00.000001Z", false); FbDump(o, "S0"); FbPushDump(o, "S0", push);
                round(new List<FbDoc>(), "2026-11-20T11:00:00Z", true); FbDump(o, "S1"); FbPushDump(o, "S1", push);
                round(docs("W3siUGMiOiJPVFJPUEMiLCJEZWxldGVkIjp0cnVlLCJVaWQiOiJ1LTAwMSIsIklkIjoiQ0FOSi0yMDI2LTAwMSIsIkJ5IjoiT1RSQSIsIlQiOiJjYW5qZSIsIkRhdG9zIjoie1wibmFtZVwiOlwiTU9IQU1NQURJTkUgQk9VVEVSRkFTU1wiLFwiZG9jXCI6XCJaNDgwMTAwM0JcIixcInN0YXR1c1wiOlwiRmluYWxpemFkb1wifSIsIlVwZGF0ZWQiOiIyMDI2LTExLTIyVDEwOjAwOjAwIn0seyJQYyI6IkRFU0tUT1AtVVRBSzNUQiIsIkRlbGV0ZWQiOmZhbHNlLCJVaWQiOiJ1LTA5MCIsIklkIjoiQ0FOSi0yMDI2LTA5MCIsIkJ5IjoiWU8iLCJUIjoiY2FuamUiLCJEYXRvcyI6IntcIm5hbWVcIjpcIkRFIEVTVEUgUENcIn0iLCJVcGRhdGVkIjoiMjAyNi0xMS0yMlQxMDowMDowMCJ9LHsiUGMiOiJPVFJPUEMiLCJEZWxldGVkIjpmYWxzZSwiVWlkIjoiaGFuZC1jYW5qZS1hYmMiLCJJZCI6IiIsIkJ5IjoiVklFSk8iLCJUIjoiY2FuamUiLCJEYXRvcyI6IntcIm5hbWVcIjpcIkFMR1VJRU4gQSBNQU5PXCJ9IiwiVXBkYXRlZCI6IjIwMjYtMTEtMTlUMDA6MDA6MDAifV0="), "2026-11-22T11:00:00Z", false); FbDump(o, "S2"); FbPushDump(o, "S2", push);
                round(new List<FbDoc>(), "", false); FbDump(o, "S3"); FbPushDump(o, "S3", push);
                w.ToastHook = keepToast5; Fb.Notify = null;
            }
            SelfTest7.Run(w, o, testDir);    // etapa 7 (sección 0)
            File.WriteAllText(Path.Combine(testDir, "cs_summary.txt"), o.ToString(), new UTF8Encoding(false));
            // aparte del resumen: al abrir GESTEM no se quita ni cambia ningún cliente con uid, tampoco los adoptados de la pestaña
            // hecha a mano (mano-…, by=Excel, su nombre sigue en la pestaña). Antes una «purga de repetidos» los borraba.
            // Sobre una copia (test\dups) y con su propio firebase.json: test\cs y cs_firebase.json se comparan después con PowerShell.
            StringBuilder dupOut = new StringBuilder(); bool kok = true;
            string keepFbPath = FbState.Path; FbState.Path = Path.Combine(dd, "firebase.json");
            Config.Cur.ExcelPath = Path.Combine(dd, "x.xlsx");
            int adopted = 0, oldVictims = 0;
            FbState.Applying = true;    // preparar: como al unirse (nada queda por enviar)
            foreach (Tramite k_t in Tramite.All)
            {
                StoreData k_s = Store.Read(k_t); if (k_s == null) continue;
                ClientRec k_h = null; int k_max = 0;
                foreach (ClientRec k_c in k_s.Clients) { if (k_c.Row >= Hand.RowBase) { if (k_h == null || k_c.Row < k_h.Row) k_h = k_c; } else if (k_c.Row > k_max) k_max = k_c.Row; }
                if (k_h != null)
                {
                    // adoptar el 1.º de la pestaña a mano como hace la unión con Firebase (FbSync): nº libre, uid mano-…, by sigue «Excel»
                    string k_nm = Ps.Str(Store.Get(k_h.F, "name"));
                    k_h.Row = k_max + 1; k_h.Id = Store.RowId(k_t, k_s.Meta, k_h.Row); k_h.Uid = Data.HandDelUid(k_t, Hand.HandKey(k_nm), "mano-"); k_h.By = "Excel";
                    k_h.Sync = true; k_h.New = false; k_h.H = null;
                    if (!k_s.Adopted.Contains(k_nm)) k_s.Adopted.Add(k_nm);
                    adopted++;
                }
                Store.Write(k_t, k_s);    // los clientes de GESTEM sin uid reciben uno
                // los que la purga quitaba: de GESTEM, by=Excel, sin borrar, con el nombre de la pestaña a mano
                HashSet<string> k_hk = new HashSet<string>();
                foreach (ClientRec k_c in k_s.Clients) if (k_c.Row >= Hand.RowBase) k_hk.Add(Hand.HandKey(Ps.Str(Store.Get(k_c.F, "name"))));
                foreach (string k_x in k_s.Adopted) k_hk.Add(Hand.HandKey(k_x));
                foreach (ClientRec k_c in k_s.Clients) if (k_c.Row < Hand.RowBase && k_c.By == "Excel" && k_c.Del == null && k_hk.Contains(Hand.HandKey(Ps.Str(Store.Get(k_c.F, "name"))))) oldVictims++;
            }
            FbState.Applying = false;
            Func<string> snap = delegate
            {
                List<string> k_l = new List<string>();
                foreach (Tramite k_t in Tramite.All)
                {
                    StoreData k_s = Store.Read(k_t); if (k_s == null) continue;
                    foreach (ClientRec k_c in k_s.Clients)
                    {
                        if (string.IsNullOrEmpty(k_c.Uid)) continue;
                        List<string> k_v = new List<string>(); foreach (Field k_f in k_t.Fields) k_v.Add(Ps.Str(Store.ToStoreValue(k_f, Store.Get(k_c.F, k_f.Key))));
                        k_l.Add(k_t.Id + "|" + k_c.Uid + "|" + k_c.Row + "|" + k_c.Id + "|" + k_c.By + "|" + k_c.Updated + "|" + (k_c.Del != null) + "|" + string.Join("|", k_v.ToArray()));
                    }
                }
                k_l.Sort(StringComparer.Ordinal); return string.Join("\n", k_l.ToArray());
            };
            string s0 = snap(); int nUid = s0.Length == 0 ? 0 : s0.Split('\n').Length, nMano = 0, nManoExcel = 0;
            foreach (string k_x in s0.Split('\n')) { string[] kq = k_x.Split('|'); if (kq.Length > 4 && kq[1].StartsWith("mano-")) { nMano++; if (kq[4] == "Excel") nManoExcel++; } }
            dupOut.Append("KEEP antes: con uid=" + nUid + " adoptados mano- (by=Excel)=" + nManoExcel + "/" + nMano + " que la purga quitaba=" + oldVictims + "\r\n");
            kok &= nManoExcel > 0 && oldVictims > 0;    // si no, la prueba no prueba nada
            for (int k_k = 1; k_k <= 2; k_k++)
            {
                // lo que hace la ventana al abrir con los datos (MainWin: SyncTabs y LoadData)
                Hand.SyncTabs(true); Data.LoadData(true, true);
                string k_s1 = snap(); bool k_same = k_s1 == s0;
                if (!k_same)
                {
                    List<string> k_a0 = new List<string>(s0.Split('\n')), k_a1 = new List<string>(k_s1.Split('\n'));
                    foreach (string k_x in k_a0) if (!k_a1.Contains(k_x)) dupOut.Append("  cambia/falta: " + k_x.Substring(0, Math.Min(160, k_x.Length)) + "\r\n");
                }
                dupOut.Append("KEEP al abrir " + k_k + ": " + (k_same ? "iguales" : "CAMBIAN") + "\r\n"); kok &= k_same;
            }
            dupOut.Append(kok ? "KEEP OK\r\n" : "KEEP MAL\r\n");
            FbState.Path = keepFbPath;
            // comprobaciones solo de C# (no hay nada igual en PowerShell): una línea «OK nombre» o «MAL nombre: …» -> result.txt
            StringBuilder only = new StringBuilder();
            {
                // config.json: las claves que esta versión no conoce se guardan tal cual (al volver a PowerShell no se pierde nada)
                Config keepCfg = Config.Cur; string keepCfgPath = App.CfgPath;
                App.CfgPath = Path.Combine(testDir, "config_claves.json");
                File.WriteAllText(App.CfgPath, "{\"Language\":\"fr\",\"ZzNueva\":{\"a\":[1,\"x\"],\"b\":null},\"zzFlag\":true,\"LastUser\":\"ANA\",\"ExcelPath\":\"C:\\\\x.xlsx\"}", new UTF8Encoding(true));
                Config.Load(); Config.Cur.LastUser = "LUIS"; Config.Save(); Config.Load();
                object j = Json.Parse(File.ReadAllText(App.CfgPath, Encoding.UTF8));
                string a = Json.WriteCompact(Json.Get(j, "ZzNueva")), fl = Ps.Str(Json.Get(j, "zzFlag"));
                bool ok = a == "{\"a\":[1,\"x\"],\"b\":null}" && fl == "True" && Config.Cur.LastUser == "LUIS" && Config.Cur.Language == "fr" && Config.Cur.ExcelPath == "C:\\x.xlsx";
                only.Append(ok ? "OK config.json conserva las claves desconocidas\r\n" : "MAL config.json conserva las claves desconocidas: ZzNueva=" + a + " zzFlag=" + fl + " user=" + Config.Cur.LastUser + "\r\n");
                App.CfgPath = keepCfgPath; Config.Cur = keepCfg; Tx.Lang = keepCfg.Language;
            }
            {
                // PDF impreso de verdad con «Microsoft Print to PDF»: la lista de clientes de SelfTest7 (apaisada, 4 páginas); el archivo es un PDF
                string pf = Path.Combine(testDir, "lista_real.pdf"); string why = "";
                try { if (File.Exists(pf)) File.Delete(pf); if (!Pdf.HasPrinter()) why = "falta la impresora"; else Pdf.Print(pf, true, MainWin.DrawClientList); }
                catch (Exception e) { why = e.Message; }
                if (why.Length == 0) { byte[] pdfHd = new byte[4]; using (FileStream fs = File.OpenRead(pf)) fs.Read(pdfHd, 0, 4); if (Encoding.ASCII.GetString(pdfHd) != "%PDF") why = "no empieza por %PDF"; }
                only.Append(why.Length == 0 ? "OK PDF impreso de verdad (Microsoft Print to PDF)\r\n" : "MAL PDF impreso de verdad: " + why + "\r\n");
            }
            {
                // etapa 7: GESTEM.exe se actualiza solo (GitHub de mentira, carpeta de instalación de prueba) y vuelve a PowerShell con «exe»: "no"
                string ed = Path.Combine(testDir, "cs_exeupd"); if (Directory.Exists(ed)) Directory.Delete(ed, true); Directory.CreateDirectory(ed);
                string exe = Path.Combine(ed, "GESTEM.exe"), oldExe = Path.Combine(ed, "GESTEM.old.exe");
                byte[] good = File.ReadAllBytes(Path.Combine(App.AppDir, "GESTEM.exe")), where = File.ReadAllBytes(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "where.exe"));
                byte[] viejo = Encoding.ASCII.GetBytes("GESTEM.exe viejo"); File.WriteAllBytes(exe, viejo); File.WriteAllText(Path.Combine(ed, "GESTEM.ps1"), "# instalado");
                byte[] serve = good; string sha = null;
                Func<byte[], string> hash = delegate(byte[] b) { using (System.Security.Cryptography.SHA256 sh2 = System.Security.Cryptography.SHA256.Create()) return BitConverter.ToString(sh2.ComputeHash(b)).Replace("-", ""); };
                Updater.FetchBytesHook = delegate(string n) { if (n != "app/GESTEM.exe") throw new RepoException(404, n); return serve; };
                Updater.FetchHook = delegate(string n) { if (n != "app/GESTEM.exe.sha256") throw new RepoException(404, n); return (sha ?? hash(serve)) + "  GESTEM.exe\n"; };
                List<string> started = new List<string>(); ExeSwitch.StartHook = delegate(string f, string a) { started.Add(Path.GetFileName(f) + " " + a); };
                string keepDir = ExeSwitch.Dir; ExeSwitch.Dir = ed;
                Func<string, string> same = delegate(string f) { return File.Exists(f) ? hash(File.ReadAllBytes(f)) : "-"; };
                Func<Action, string> err = delegate(Action a) { try { a(); return ""; } catch (Exception e) { return e.Message; } };
                try
                {
                    string e1 = err(delegate { ExeSwitch.Update("2.0.1"); });
                    bool ok1 = e1 == "" && same(exe) == hash(good) && same(oldExe) == hash(viejo) && !File.Exists(Path.Combine(ed, "GESTEM.new.exe"));
                    only.Append(ok1 ? "OK actualización de GESTEM.exe (normal, el nuevo en su sitio, el anterior en GESTEM.old.exe)\r\n" : "MAL actualización de GESTEM.exe (normal): " + e1 + " exe=" + same(exe) + " old=" + same(oldExe) + "\r\n");
                    File.WriteAllBytes(exe, viejo); File.Delete(oldExe);
                    sha = new string('0', 64); string e2 = err(delegate { ExeSwitch.Update("2.0.2"); }); sha = null;
                    serve = where; string e3 = err(delegate { ExeSwitch.Update("2.0.3"); }); serve = good;
                    bool ok2 = e2.StartsWith("SHA-256 no coincide") && e3.StartsWith("--smoke falla") && same(exe) == hash(viejo) && !File.Exists(oldExe) && !File.Exists(Path.Combine(ed, "GESTEM.new.exe"));
                    only.Append(ok2 ? "OK actualización de GESTEM.exe (SHA-256 malo y --smoke que falla, GESTEM.exe sin tocar)\r\n" : "MAL actualización de GESTEM.exe (fallos): [" + e2 + "] [" + e3 + "] exe=" + same(exe) + "\r\n");
                    ExeSwitch.Restart();
                    string me = System.Diagnostics.Process.GetCurrentProcess().Id.ToString();
                    bool ok3 = started.Count == 1 && started[0].StartsWith("GESTEM.exe ") && started[0].EndsWith(" --tras " + me) && !ExeSwitch.CanUpdate;
                    only.Append(ok3 ? "OK actualización de GESTEM.exe (se abre el nuevo con --tras; en modo prueba no se actualiza sola)\r\n" : "MAL actualización de GESTEM.exe (abrir el nuevo): " + string.Join(" / ", started.ToArray()) + " CanUpdate=" + ExeSwitch.CanUpdate + "\r\n");
                    // «exe»: "no" solo con la copia instalada (GESTEM.ps1 al lado y exe_instalado.txt); abre GESTEM.ps1 -VolverRemoto
                    started.Clear(); string done = Path.Combine(App.DataDir, "exe_instalado.txt"); if (File.Exists(done)) File.Delete(done);
                    bool w0 = ExeSwitch.WantsPs(Json.Parse("{\"exe\":\"no\"}"));
                    File.WriteAllText(done, "{}");
                    bool w1 = ExeSwitch.WantsPs(Json.Parse("{\"exe\":\"no\"}")), w2 = ExeSwitch.WantsPs(Json.Parse("{\"exePcs\":\"*\"}")), w3 = ExeSwitch.WantsPs(null);
                    File.Delete(done);
                    ExeSwitch.BackToPs();
                    bool ok4 = !w0 && w1 && !w2 && !w3 && started.Count == 1 && started[0] == "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File \"" + Path.Combine(ed, "GESTEM.ps1") + "\" -Oculto -VolverRemoto";
                    only.Append(ok4 ? "OK vuelta a PowerShell con control.json «exe» = \"no\" (solo la copia instalada)\r\n" : "MAL vuelta a PowerShell: " + w0 + w1 + w2 + w3 + " " + string.Join(" / ", started.ToArray()) + "\r\n");
                }
                finally { Updater.FetchBytesHook = null; Updater.FetchHook = null; ExeSwitch.StartHook = null; ExeSwitch.Dir = keepDir; }
            }
            // modo prueba: fuera de las carpetas de prueba no se escribe nada (sin --real)
            StringBuilder gd = new StringBuilder(); bool gok = true;
            string appd = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GESTEM");
            foreach (string rp in new string[] { Path.Combine(appd, "config.json"), Path.Combine(appd, "gestem.log"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "GESTEM_MANAGER_1", "GESTEM MANAGER", "GESTEM datos"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "GESTEM_MANAGER_1", "GESTEM MANAGER", "GESTEM documentos") })
            {
                bool blocked = false; try { App.Guard(rp, "prueba"); } catch (GestemException) { blocked = true; }
                gd.Append("GUARD " + (blocked ? "parado " : "PERMITIDO ") + rp + "\r\n"); gok &= blocked;
            }
            gd.Append("GUARD " + (App.IsSafe(testDir) ? "permitido " : "PARADO ") + "carpeta de prueba\r\n"); gok &= App.IsSafe(testDir);
            // arreglar carpetas en unos documentos que no son de prueba: no se mueve nada
            string outside = Path.Combine(Path.GetTempPath(), "gestem_guard_" + Guid.NewGuid().ToString("N"));
            string mis = Path.Combine(outside, "DGT ACTUALIZADO", ids[2] + " - " + Docs.SafeName(RowById(ids[1]).Name));
            Directory.CreateDirectory(mis); File.WriteAllText(Path.Combine(mis, "x.pdf"), "X");
            Config.Cur.DocsDir = outside;
            Docs.RepairAll();
            bool still = Directory.Exists(mis) && File.Exists(Path.Combine(mis, "x.pdf"));
            gd.Append("GUARD arreglar documentos fuera de prueba: " + (still ? "no tocados" : "MOVIDOS") + "\r\n"); gok &= still;
            try { Directory.Delete(outside, true); } catch { }
            Config.Cur.DocsDir = droot;
            gd.Append(gok ? "GUARD OK\r\n" : "GUARD MAL\r\n");
            File.WriteAllText(Path.Combine(testDir, "cs_guard.txt"), gd.ToString(), new UTF8Encoding(false));            File.WriteAllText(Path.Combine(testDir, "cs_keep.txt"), dupOut.ToString(), new UTF8Encoding(false)); File.WriteAllText(Path.Combine(testDir, "cs_only.txt"), only.ToString(), new UTF8Encoding(false));
            return 0;
        }

        // Get-TestRow de pstest.ps1: cliente de GESTEM sin NIE repetido
        static ClientRow TestRow(int skip)
        {
            int k = 0;
            foreach (ClientRow r in G.Rows)
            {
                if (r.Row >= Hand.RowBase) continue;
                string dc = Ps.Str(r.Doc).Trim();
                if (dc.Length > 0) { int n = 0; foreach (ClientRow x in G.Rows) if (string.Equals(Ps.Str(x.Doc).Trim(), dc, StringComparison.OrdinalIgnoreCase)) n++; if (n != 1) continue; }
                if (k == skip) return r; k++;
            }
            return null;
        }
        static ClientRow RowAt(int row) { foreach (ClientRow r in G.Rows) if (r.Row == row) return r; return null; }
        static List<ClientRow> HandRows() { List<ClientRow> l = new List<ClientRow>(); foreach (ClientRow r in G.Rows) if (r.Row >= Hand.RowBase) l.Add(r); return l; }
        static string RowText(string tag, int row)
        {
            ClientRow r = RowAt(row);
            if (r == null) return tag + " " + row + " null";
            return tag + " " + row + " " + r.Id + "|" + r.Name + "|" + r.Status + "|" + r.Pay + "|" + N(r.Pending) + "|" + r.Resp + " n=" + G.Rows.Count;
        }
        static ClientRow RowById(string id) { foreach (ClientRow r in G.Rows) if (r.Id == id) return r; return null; }
        static void DocsTree(StringBuilder o, string tag, string root)
        {
            List<string> l = new List<string>();
            foreach (string e in Directory.GetFileSystemEntries(root, "*", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(e) == "clientes.txt") continue;
                l.Add(e.Substring(root.Length + 1) + (File.Exists(e) ? " " + new FileInfo(e).Length : ""));
            }
            l.Sort(StringComparer.Ordinal);
            foreach (string x in l) o.Append(tag + " " + x + "\r\n");
        }
        static List<string> Items(ComboBox c) { List<string> l = new List<string>(); foreach (object x in c.Items) l.Add(Ps.Str(x)); return l; }
        static string StatsText(MainWin w)
        {
            List<string> l = new List<string>();
            foreach (UIElement ch in w.U<Grid>("StatsGrid").Children)
            {
                int r = Grid.GetRow(ch), cl = Grid.GetColumn(ch);
                if (ch is TextBlock) l.Add(r + "," + cl + "=" + ((TextBlock)ch).Text);
                else if (ch is StackPanel) { StackPanel sp = (StackPanel)ch; l.Add(r + ",n=" + ((TextBlock)sp.Children[1]).Text + "|" + ((System.Windows.Shapes.Ellipse)sp.Children[0]).Fill); }
                else if (ch is Grid) l.Add(r + ",bar=" + ((Grid)ch).ColumnDefinitions[0].Width.Value.ToString("R", Ps.Inv));
                else if (ch is FrameworkElement && ((FrameworkElement)ch).Tag != null) l.Add(r + ",hit=" + ((FrameworkElement)ch).Tag);
            }
            List<string> lab = new List<string>();
            foreach (UIElement ch in w.U<Grid>("ChartLabels").Children) { StackPanel sp = (StackPanel)ch; lab.Add(((TextBlock)sp.Children[0]).Text + "/" + ((TextBlock)sp.Children[1]).Text); }
            List<string> bk = new List<string>();
            foreach (ChartBucket b in w.ChartBuckets) bk.Add(b.N + ":" + b.Paid.ToString("R", Ps.Inv));
            Func<string, string> T = delegate(string n) { return w.U<TextBlock>(n).Text; };
            return "sub=" + T("StatsSub") + "|table=" + T("StatsTable") + "|all=" + w.U<FrameworkElement>("StatsAllBtn").Visibility + "|kpi=" + T("KpiBill") + ";" + T("KpiBillF") + ";" + T("KpiPaid") + ";" + T("KpiPaidF") + ";" + T("KpiPend") + ";" + T("KpiPendF") + ";" + T("KpiNet") + ";" + T("KpiNetF") + ";" + w.U<Grid>("KpiPaidTrack").ColumnDefinitions[0].Width.Value.ToString("R", Ps.Inv) + "\r\n  "
                + string.Join(" ; ", l) + "\r\n  chart=" + T("ChartTitle") + "|" + T("ChartSub") + "|" + T("ChartDelta") + "|" + string.Join(" ", lab) + "|" + string.Join(" ", bk);
        }        static string SetText(MainWin w)
        {
            Func<string, string> T = delegate(string n) { return ((TextBox)w.U<FrameworkElement>(n)).Text; };
            System.Windows.Media.ImageBrush ib = w.U<System.Windows.Shapes.Shape>("LogoPreview").Fill as System.Windows.Media.ImageBrush;
            string lg = ib != null && ib.ImageSource is System.Windows.Media.Imaging.BitmapImage ? "propio" : "de serie";
            ComboBox lb = w.U<ComboBox>("LangBox"), nb = w.U<ComboBox>("NatBox"); TextBlock el = w.U<TextBlock>("ExtLastTxt");
            return "file=" + Path.GetFileName(T("FilePathBox")) + "|docs=" + System.Text.RegularExpressions.Regex.Replace(Path.GetFileName(T("DocsDirBox")), "^cs_", "X_") + "|lang=" + lb.SelectedIndex + "/" + lb.Items.Count + "|nat=" + nb.Text + "/" + nb.Items.Count + "|col=" + T("ColleaguesBox") + "|co=" + T("CompanyBox")
                + "|inv=" + T("InvNameBox") + ";" + T("InvActBox") + ";" + T("InvNifBox") + ";" + T("InvIbanBox") + ";" + T("InvDomBox") + ";" + T("InvPobBox") + ";" + T("InvNextBox") + "|tasas=" + T("TasasBox")
                + "|ext=" + T("ExtDirBox") + "|extLast=" + el.Text + ";" + el.Foreground + ";" + w.U<Button>("ExtOpenBtn").IsEnabled + "|bk=" + w.U<CheckBox>("BackupChk").IsChecked + "|open=" + w.U<CheckBox>("OpenAfterChk").IsChecked + "|welcome=" + w.U<FrameworkElement>("WelcomeText").Visibility + "|logo=" + lg;
        }
        static string CfgText(MainWin w)
        {
            Config f = Config.Cur; string[] lc; G.Lists.TryGetValue("colleagues", out lc);
            Func<string, string> T = delegate(string n) { return w.U<TextBlock>(n).Text; };
            return "lang=" + f.Language + "|nat=" + f.DefaultNationality + "|co=" + f.CompanyName + "|inv=" + f.InvName + ";" + f.InvActivity + ";" + f.InvNif + ";" + f.InvIban + ";" + f.InvAddress + ";" + f.InvTown + "|tasas=" + f.TasasCanje
                + "|col=" + string.Join(",", f.Colleagues) + "|lists=" + (lc == null ? "" : string.Join(",", lc))
                + "|bk=" + f.Backup + "|open=" + f.OpenAfterSave + "|ext=" + f.ExtBackupDir + "|extLast=" + f.LastExtBackup + "|docs=" + System.Text.RegularExpressions.Regex.Replace(Path.GetFileName(f.DocsDir), "^cs_", "X_") + "|logo=" + Path.GetFileName(f.CompanyLogo) + "|logoFile=" + File.Exists(f.CompanyLogo)
                + "|title=" + w.Win.Title.Replace("  —  " + TestMode.Banner, "") + "|side=" + T("SideCompany") + "/" + w.U<FrameworkElement>("SideCompany").Visibility + "|sess=" + T("SessCompany") + "|about=" + T("AboutCompany") + "|status=" + T("StatusText");
        }        static string SessText(MainWin w)
        {
            List<string> nm = new List<string>();
            foreach (UIElement ch in w.U<Panel>("SessNames").Children) { Button b = (Button)ch; nm.Add(Ps.Str(b.Content) + (b.FontWeight == FontWeights.SemiBold ? "*" : "")); }
            string[] lc; G.Lists.TryGetValue("colleagues", out lc);
            return "names=" + string.Join(",", nm) + "|vis=" + w.U<FrameworkElement>("PanelSession").Visibility + "|hint=" + w.U<FrameworkElement>("SessNewHint").Visibility + "|user=" + App.User + "|last=" + Config.Cur.LastUser
                + "|col=" + string.Join(",", Config.Cur.Colleagues) + "|lists=" + (lc == null ? "" : string.Join(",", lc)) + "|btn=" + Ps.Str(w.U<ContentControl>("UserBtn").Content) + "|title=" + w.U<TextBlock>("SessTitle").Text;
        }        static string ColState(MainWin w)
        {
            Func<string, string> T = delegate(string n) { return w.U<TextBlock>(n).Text; };
            List<string> li = new List<string>(); foreach (ColabListItem x in (System.Collections.IEnumerable)w.U<ListBox>("ColabList").ItemsSource ?? new ColabListItem[0]) li.Add(x.Key + "~" + x.Name + "~" + (int)x.Glyph[0] + "~" + x.Info);
            List<string> gi = new List<string>(); foreach (ColabItem x in (System.Collections.IEnumerable)w.U<DataGrid>("ColabGrid").ItemsSource ?? new ColabItem[0])
                gi.Add(x.TramId + "|" + x.Tram + "|" + x.Row + "|" + x.Id + "|" + x.Name + "|" + x.Colab + "|" + x.FechaTxt + "|" + x.Status + "|" + x.Color + "|" + x.PriceTxt + "|" + x.PaidTxt + "|" + x.PendTxt + "|" + x.ComTxt + "|" + x.ComStTxt + "|" + x.ComColor);
            return "name=" + T("ColabName") + "|k=" + T("ColK1") + ";" + T("ColK2") + ";" + T("ColK3") + ";" + T("ColK4") + ";" + T("ColK5") + ";" + T("ColK6") + "|empty=" + w.U<FrameworkElement>("ColabEmpty").Visibility + ";" + T("ColabEmpty")
                + "|grid=" + w.U<FrameworkElement>("ColabGrid").Visibility + "|list=" + string.Join(",", li) + "|sel=" + w.U<ListBox>("ColabList").SelectedIndex + "\r\n  " + string.Join("\r\n  ", gi);
        }        static string RemState(MainWin w)
        {
            List<string> f = new List<string>();
            foreach (string n in new string[] { "control.cache", "control.visto" }) { string p = Path.Combine(App.DataDir, n); f.Add(n + "=" + (File.Exists(p) ? BitConverter.ToString(File.ReadAllBytes(p)) : "-")); }
            return "remote=" + (w.Remote != null ? Json.WriteCompact(w.Remote) : "-") + "|paused=" + w.U<FrameworkElement>("PanelPaused").Visibility + "|" + w.U<TextBlock>("PausedText").Text + "|" + string.Join(";", f);
        }        static void FbDump(StringBuilder o, string tag)
        {
            FbState st = FbState.Load(); Dictionary<string, int> hs = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase); int nn = 0;
            foreach (System.Collections.DictionaryEntry e in st.Pend) { PendItem p = (PendItem)e.Value; int v; hs.TryGetValue(p.H, out v); hs[p.H] = v + 1; if (p.N) nn++; }
            List<string> hk = new List<string>(hs.Keys); Ps.Sort(hk, StringComparer.CurrentCultureIgnoreCase.Compare); for (int i = 0; i < hk.Count; i++) hk[i] = hk[i] + "=" + hs[hk[i]];
            o.Append("FB " + tag + " joined=" + st.Joined + " lastMod=" + st.LastMod + " pend=" + st.Pend.Count + " h=" + string.Join(",", hk) + " nuevos=" + nn + " err=" + Fb.LastErr + "\r\n");
            foreach (string tid in new string[] { "canje", "autonomo" })
            {
                StoreData s = Store.Read(Tramite.ById(tid)); List<ClientRec> l = new List<ClientRec>();
                foreach (ClientRec c in s.Clients) if (Array.IndexOf(new string[] { "u-001", "u-002", "u-090", "u-a01" }, c.Uid) >= 0 || c.Row == 12 || c.Row == 88 || c.Row == 89 || Ps.Str(Store.Get(c.F, "name")) == "AZIZ BOUTERFASS") l.Add(c);
                Ps.Sort(l, delegate(ClientRec a, ClientRec b) { return a.Row.CompareTo(b.Row); });
                foreach (ClientRec c in l) o.Append("FBC " + tag + " " + tid + " " + c.Row + " " + c.Id + " " + System.Text.RegularExpressions.Regex.Replace(Ps.Str(c.Uid), "^[0-9a-f]{32}$", "GUID") + "|" + Ps.Str(Store.Get(c.F, "name")) + "|" + Ps.Str(Store.Get(c.F, "status")) + "|" + Ps.Str(Store.Get(c.F, "price")) + "|" + Ps.Str(Store.Get(c.F, "xI")) + "|" + c.Updated + "|" + c.By + "|sync=" + c.Sync + "|del=" + (c.Del != null) + "\r\n");
                List<string> h = new List<string>(); foreach (HandDelRec e in s.HandDel) h.Add(e.Uid + ":" + e.Name + ":" + e.On + ":" + e.At);
                o.Append("FBH " + tag + " " + tid + " " + string.Join(",", h) + " adopted=" + s.Adopted.Count + "\r\n");
            }
        }
        static void FbPushDump(StringBuilder o, string tag, List<FbPush> push)
        {
            List<string> l = new List<string>(); int bad = 0;
            foreach (FbPush p in push) { string x = p.T + "|" + p.Id + "|" + p.Datos + "|" + p.Updated + "|" + p.By + "|" + p.H + "|" + p.N + "|" + p.Deleted + "|" + p.Uid; l.Add(x); foreach (Field f in Tramite.ById(p.T).Fields) if (f.Local && p.Datos.Contains("\"" + f.Key + "\":")) { bad++; break; } }
            for (int i = 0; i < l.Count; i++) l[i] = System.Text.RegularExpressions.Regex.Replace(l[i], "[0-9a-f]{32}", "GUID"); l.Sort(StringComparer.Ordinal);    // el orden sale de un @{} con un GUID al azar: no cuenta
            o.Append("FBP " + tag + " n=" + l.Count + " conContraseña=" + bad + "\r\n"); int k = 0;
            foreach (string x in l) if ((x.StartsWith("canje|") || x.Contains("|u-")) && k++ < 12) o.Append("FBP " + tag + " " + System.Text.RegularExpressions.Regex.Replace(x, "[0-9a-f]{32}", "GUID") + "\r\n");
        }        static void AdmTree(object node, List<string> l)
        {
            DependencyObject d = node as DependencyObject; if (d == null) return;
            if (d is TextBlock) l.Add(((TextBlock)d).Text); else if (d is ContentControl && ((ContentControl)d).Content is string) l.Add((string)((ContentControl)d).Content);
            foreach (object ch in LogicalTreeHelper.GetChildren(d)) AdmTree(ch, l);
        }
        static string AdmState(MainWin w)
        {
            Func<string, string> V = delegate(string n) { return w.U<FrameworkElement>(n).Visibility.ToString(); };
            Func<string, string> T = delegate(string n) { return w.U<TextBlock>(n).Text; };
            List<string> sec = new List<string>(); foreach (string n in new string[] { "SecFileBox", "SecExtBox", "SecCompBox", "SecTeamBox", "SecPrefBox", "SecHelpBox", "NavSettings" }) sec.Add(V(n));
            return "admin=" + App.IsAdmin + "|login=" + V("PanelAdminLogin") + "|panel=" + V("PanelAdmin") + "|pw2=" + V("AdmPw2") + "|hint=" + T("AdmLoginHint") + "|msg=" + T("AdmLoginMsg") + "|locksMsg=" + T("AdmLocksMsg")
                + "|pwMsg=" + T("AdmPwMsg") + ";" + w.U<TextBlock>("AdmPwMsg").Foreground + "|fb=" + T("AdmFbState") + "|sec=" + string.Join(",", sec) + "|file=" + File.Exists(MainWin.AdminPath);
        }        static string DocsPanelText(MainWin w)
        {
            List<string> it = new List<string>();
            foreach (DocItem d in (System.Collections.IEnumerable)w.U<ListBox>("DocsList").ItemsSource ?? new DocItem[0]) it.Add(d.Name + "|" + d.Info + "|" + (int)d.Glyph[0]);
            return "title=" + w.U<TextBlock>("DocsTitle").Text + "|sub=" + w.U<TextBlock>("DocsSub").Text + "|msg=" + w.U<TextBlock>("DocsMsg").Text + "|count=" + w.U<TextBlock>("DocsCount").Text
                + "|empty=" + w.U<FrameworkElement>("DocsEmpty").Visibility + "|vis=" + w.U<FrameworkElement>("PanelDocs").Visibility + "|" + string.Join(" ; ", it.ToArray());
        }
        static string FormLine(string head, Tramite t, Dictionary<string, object> v)
        {
            StringBuilder line = new StringBuilder(head);
            foreach (Field f in t.Fields) line.Append("|" + f.Key + "=" + Ps.Str(Store.ToStoreValue(f, Store.Get(v, f.Key))));
            return line.ToString();
        }
        static string Card(MainWin w)
        {
            Func<string, string> T = delegate(string n) { return w.U<System.Windows.Controls.TextBlock>(n).Text; };
            Func<string, bool> Vis = delegate(string n) { return w.U<System.Windows.FrameworkElement>(n).Visibility == System.Windows.Visibility.Visible; };
            return "id=" + T("CardId") + "|tag=" + T("CardTag") + "|name=" + T("CardName") + "|sub=" + T("CardSub") + "|tipo=" + T("CardTipo")
                + "|st=" + (Vis("CardStatusB") ? T("CardStatus") : "") + "|price=" + T("SumPrice") + "|iva=" + (Vis("SumIva") ? T("SumIva") : "")
                + "|mg=" + (Vis("SumMargin") ? T("SumMargin") : "") + "|paid=" + T("SumPaid") + "|pend=" + T("SumPend") + "|title=" + T("FormTitle");
        }
    }
}

// Prueba de igualdad con la versión PowerShell: lee los datos de una carpeta de prueba, calcula cada fila
// y vuelve a guardar los archivos. tools\pstest.ps1 hace lo mismo con las funciones de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.IO;
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
            App.DataDir = testDir; App.BackupDir = Path.Combine(testDir, "copias"); App.LogPath = Path.Combine(testDir, "cs_log.txt"); App.CfgPath = Path.Combine(testDir, "config_cs.json"); FbState.Path = Path.Combine(testDir, "cs_firebase.json");
            Directory.CreateDirectory(App.BackupDir);
            Config.Cur = new Config();
            Config.Cur.ExcelPath = Path.Combine(testDir, "cs", "x.xlsx");
            // copia de los datos tal como llegan, para comprobar al final que no quedan clientes repetidos (test\dups)
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
            File.WriteAllText(Path.Combine(testDir, "cs_summary.txt"), o.ToString(), new UTF8Encoding(false));
            // aparte del resumen (PowerShell 1.9.2 sí deja los repetidos): sin clientes repetidos después de cargar
            System.Text.StringBuilder dupOut = new System.Text.StringBuilder();
            int bad = 0;
            // sobre una copia (test\dups): los datos de test\cs se comparan después con los de PowerShell
            Config.Cur.ExcelPath = Path.Combine(dd, "x.xlsx");
            int before = 0;
            foreach (Tramite t in Tramite.All) { StoreData s0 = Store.Read(t); if (s0 != null) before += s0.Clients.Count; }
            Hand.SyncTabs(true, true);
            dupOut.Append("DUP antes: clientes=" + before + "\r\n");
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = Store.Read(t); if (s == null) continue;
                HashSet<string> gest = new HashSet<string>(); int dup = 0;
                foreach (ClientRec c in s.Clients) if (c.Row < Hand.RowBase && c.Del == null) gest.Add(Hand.HandKey(Ps.Str(Store.Get(c.F, "name"))));
                foreach (ClientRec c in s.Clients) if (c.Row >= Hand.RowBase && gest.Contains(Hand.HandKey(Ps.Str(Store.Get(c.F, "name"))))) dup++;
                bad += dup;
                dupOut.Append("DUP " + t.Id + " clientes=" + s.Clients.Count + " repetidos=" + dup + "\r\n");
            }
            dupOut.Append(bad == 0 ? "DUP OK\r\n" : "DUP MAL\r\n");
            File.WriteAllText(Path.Combine(testDir, "cs_dups.txt"), dupOut.ToString(), new UTF8Encoding(false));
            return 0;
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

// Excel por COM: se usa cuando el archivo está abierto en Excel o la escritura directa no es segura.
// Copiado de Invoke-Excel / Open-Book / Close-Book / Find-Sheet / Get-Sheet / New-TramiteSheet / Read-TramiteData /
// Restore-Formula / Set-CellValue / Save-Client.
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Gestem
{
    public class Book { public object Xl; public object Wb; public bool Mine; }

    // llamadas a Excel por IDispatch (como PowerShell); sin «dynamic», que falla al cerrar y volver a abrir Excel
    public static class Com
    {
        static object Inv(object o, string n, System.Reflection.BindingFlags f, object[] a)
        {
            try { return o.GetType().InvokeMember(n, f, null, o, a); }
            catch (System.Reflection.TargetInvocationException e)
            {
                if (e.InnerException != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw();
                throw;
            }
        }
        public static object Get(object o, string n, params object[] a) { return Inv(o, n, System.Reflection.BindingFlags.GetProperty, a); }
        public static void Set(object o, string n, object v) { Inv(o, n, System.Reflection.BindingFlags.SetProperty, new object[] { v }); }
        public static object Call(object o, string n, params object[] a) { return Inv(o, n, System.Reflection.BindingFlags.InvokeMethod, a); }
        public static object Range(object ws, string addr) { return Get(ws, "Range", addr); }
        public static object V2(object ws, string addr) { return Get(Range(ws, addr), "Value2"); }
        public static string S(object ws, string addr) { return Ps.Str(V2(ws, addr)); }
        public static void SetV2(object ws, string addr, object v) { Set(Range(ws, addr), "Value2", v); }
        public static List<object> Items(object coll)
        {
            List<object> l = new List<object>();
            int n = Convert.ToInt32(Get(coll, "Count"));
            for (int i = 1; i <= n; i++) l.Add(Get(coll, "Item", i));
            return l;
        }
        // celdas de un rango, también si tiene varias áreas (como foreach en PowerShell)
        public static List<object> Cells(object rng)
        {
            List<object> l = new List<object>();
            foreach (object area in Items(Get(rng, "Areas")))
            {
                object cells = Get(area, "Cells");
                int n = Convert.ToInt32(Get(cells, "Count"));
                for (int i = 1; i <= n; i++) l.Add(Get(cells, "Item", i));
            }
            return l;
        }
    }

    // estado que en PowerShell era $script:… (trámite actual, listas, caché…)
    public static class G
    {
        public static Tramite T;
        public static Dictionary<string, double> Tariff = new Dictionary<string, double>();
        public static Dictionary<string, string[]> Lists = new Dictionary<string, string[]>();
        public static double? DefaultPrice;
        public static List<ClientRow> Rows = new List<ClientRow>();
        public static int? FirstEmpty;
        public static string NextId = "";
        public static Dictionary<string, TramiteData> Cache;
        public static long CacheStamp;
        public static Hashtable MissingTabs = PsH.New();
        public static Hashtable MissingWarned = PsH.New();
        // la ventana se engancha aquí (Use-TramiteData, Set-Status, Show-Error)
        public static Action<TramiteData> OnUseData;
        public static Action<string> OnStatus;
        public static Action<string> OnError;
        public static void Status(string s) { if (OnStatus != null) OnStatus(s); }
        public static void Error(string s) { if (OnError != null) OnError(s); }
    }

    public static class ExcelCom
    {
        public static bool Installed() { return Type.GetTypeFromProgID("Excel.Application") != null; }

        public static bool IsBusy(Exception e)
        {
            Exception x = e;
            while (x != null) { if (x.HResult == -2147418111 || x.HResult == -2147417846) return true; x = x.InnerException; }
            return false;
        }
        static int ComHr(Exception e)
        {
            if (e is COMException) return e.HResult;
            if (e.InnerException is COMException) return e.InnerException.HResult;
            return 0;
        }

        // Invoke-Excel: Excel rechaza llamadas si está ocupado: reintentar un poco
        public static T Run<T>(Func<T> a)
        {
            for (int i = 0; ; i++)
            {
                try { return a(); }
                catch (Exception e)
                {
                    int hr = ComHr(e);
                    bool busy = hr == -2147418111 || hr == -2147417846;
                    if (busy && i < 5) { System.Threading.Thread.Sleep(600); continue; }
                    if (busy) throw new GestemException(Tx.L("errBusy"));
                    throw;
                }
            }
        }
        public static void Run(Action a) { Run<bool>(delegate { a(); return true; }); }

        public static void Release(object o) { if (o != null) { try { Marshal.FinalReleaseComObject(o); } catch { } } }

        public static Book Open(bool forWrite)
        {
            string path = Config.Cur.ExcelPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) throw new GestemException(Tx.L("errNoFile"));
            string full = Path.GetFullPath(path);
            string fname = Path.GetFileName(full);
            object xl = null, wb = null; bool mine = false;
            try { xl = Marshal.GetActiveObject("Excel.Application"); } catch { xl = null; }
            if (xl != null)
            {
                try
                {
                    foreach (object w in Com.Items(Com.Get(xl, "Workbooks")))
                    {
                        string fn = Ps.Str(Com.Get(w, "FullName"));
                        if (Ps.EqI(fn, full) || (PsH.Is(fn, "^https?:") && Ps.EqI(Ps.Str(Com.Get(w, "Name")), fname))) { wb = w; break; }
                    }
                }
                catch (Exception e) { if (IsBusy(e)) throw; }      // un Excel que se está cerrando: se abre otro
            }
            if (wb == null)
            {
                xl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                Com.Set(xl, "Visible", false); Com.Set(xl, "DisplayAlerts", false); Com.Set(xl, "ScreenUpdating", false);
                wb = Com.Call(Com.Get(xl, "Workbooks"), "Open", full);
                mine = true;
                if (forWrite && (bool)Com.Get(wb, "ReadOnly"))
                {
                    Com.Call(wb, "Close", false); Com.Call(xl, "Quit");
                    Release(wb); Release(xl);
                    throw new GestemException(Tx.L("errLocked"));
                }
            }
            Book b = new Book(); b.Xl = xl; b.Wb = wb; b.Mine = mine;
            return b;
        }

        public static void Close(Book h, bool save)
        {
            if (h == null) return;
            try
            {
                if (save) Com.Call(h.Wb, "Save");
                if (h.Mine) { Com.Call(h.Wb, "Close", false); Com.Call(h.Xl, "Quit"); }
            }
            finally
            {
                if (h.Mine) { Release(h.Wb); Release(h.Xl); }
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            }
        }

        static List<object> Sheets(object wb) { return Com.Items(Com.Get(wb, "Worksheets")); }
        static string Name(object w) { return Ps.Str(Com.Get(w, "Name")); }

        // Find-Sheet: la pestaña aunque tenga otra tilde, mayúsculas o espacios; null si de verdad no está
        public static object FindSheet(object wb, string name)
        {
            try { return Com.Get(Com.Get(wb, "Worksheets"), "Item", name); } catch (Exception e) { if (IsBusy(e)) throw; }
            string key = Xlsx.SheetKey(name);
            foreach (object w in Sheets(wb)) if (Xlsx.SheetKey(Name(w)) == key) return w;
            return null;
        }

        // Test-GestemLayout: cabecera «Cliente»/«Nombre…» en la columna del nombre, encima de la primera fila
        static bool GestemLayout(object ws, Tramite t)
        {
            if (!t.IsGen) return true;
            string hdr = "";
            try { hdr = Xlsx.SheetKey(Com.S(ws, t.NameCol + (t.FirstRow - 1))); } catch (Exception e) { if (IsBusy(e)) throw; return false; }
            return hdr.StartsWith("CLIENTE") || hdr.StartsWith("NOMBRE");
        }

        public static object FindTramiteSheet(object wb, Tramite t)
        {
            List<object> cands = new List<object>();
            object ws = FindSheet(wb, t.Sheet); if (ws != null) cands.Add(ws);
            if (Hand.SheetAliases.ContainsKey(t.Id))
                foreach (string a in Hand.SheetAliases[t.Id]) if (!string.IsNullOrEmpty(a)) { ws = FindSheet(wb, a); if (ws != null) cands.Add(ws); }
            string key = Xlsx.SheetKey(t.Sheet);
            foreach (object w in Sheets(wb))
            {
                string k = Xlsx.SheetKey(Name(w));
                if (k.StartsWith(key)) { string rest = k.Substring(key.Length); if (rest == "S" || rest == "ES" || rest.StartsWith(" ")) cands.Add(w); }
            }
            foreach (object c in cands) if (GestemLayout(c, t)) return c;
            return null;
        }

        public static long ExcelStamp() { try { return File.GetLastWriteTimeUtc(Config.Cur.ExcelPath).Ticks; } catch { return 0; } }

        public static bool TabKnownMissing(Tramite t)
        {
            if (!G.MissingTabs.ContainsKey(t.Id)) return false;
            if ((long)G.MissingTabs[t.Id] == ExcelStamp()) return true;
            G.MissingTabs.Remove(t.Id); G.MissingWarned.Remove(t.Id);     // el Excel cambió: se vuelve a mirar
            return false;
        }

        // Get-Sheet: con create, si al Excel le falta la pestaña del trámite, GESTEM la crea
        public static object GetSheet(Book h, bool create)
        {
            Tramite T = G.T;
            object ws = Run<object>(delegate { return FindTramiteSheet(h.Wb, T); });
            if (ws != null) return ws;
            if (create && T.IsGen) return Run<object>(delegate { return NewTramiteSheet(h, T); });
            List<string> names = new List<string>();
            try { foreach (object w in Sheets(h.Wb)) names.Add("«" + Name(w) + "»"); } catch { }
            string msg = string.Format(Tx.L("errNoSheet"), T.Sheet);
            if (names.Count > 0) msg += "\r\n\r\n" + string.Format(Tx.L("errTabsFound"), string.Join(", ", names.ToArray()));
            G.MissingTabs[T.Id] = ExcelStamp();
            throw new GestemException(msg);
        }

        public static object NewTramiteSheet(Book h, Tramite t)
        {
            object wb = h.Wb;
            // si ya hay una pestaña suya con ese nombre (hecha a mano), la de GESTEM se llama «… GESTEM»
            string name = t.Sheet;
            List<string> taken = new List<string>();
            foreach (object w in Sheets(wb)) taken.Add(Xlsx.SheetKey(Name(w)));
            if (PsH.ContainsI(taken, Xlsx.SheetKey(name))) name = t.Sheet + " GESTEM";
            object sheets = Com.Get(wb, "Worksheets");
            object last = Com.Get(sheets, "Item", Convert.ToInt32(Com.Get(sheets, "Count")));
            object ws = Com.Call(sheets, "Add", System.Reflection.Missing.Value, last);
            Com.Set(ws, "Name", name);
            Com.SetV2(ws, "A1", t.TitleEs.ToUpper());
            object a1f = Com.Get(Com.Range(ws, "A1"), "Font"); Com.Set(a1f, "Bold", true); Com.Set(a1f, "Size", 14);
            Com.SetV2(ws, "A2", "Pestaña creada por GESTEM MANAGER");
            int hr = t.FirstRow - 1;
            Hashtable heads = PsH.New();
            heads["A"] = "ID Expediente"; heads["L"] = "Pendiente (€)"; heads["T"] = "Días"; heads["W"] = "Neto (€)";
            foreach (Field f in t.Fields) { string es; f.Label.TryGetValue("es", out es); heads[f.Col] = es ?? ""; }
            foreach (DictionaryEntry k in heads) Com.SetV2(ws, (string)k.Key + hr, (string)k.Value);
            object hdr = Com.Range(ws, "A" + hr + ":" + t.LastCol + hr);
            Com.Set(Com.Get(hdr, "Font"), "Bold", true); Com.Set(Com.Get(hdr, "Interior"), "Color", 14277081);
            int f1 = t.FirstRow, l1 = t.LastRow;
            Meta meta = null; try { StoreData s = Store.Read(t); if (s != null) meta = s.Meta; } catch { }
            if (meta == null) meta = Hand.NewEmptyStore(t, null).Meta;
            int n = l1 - f1 + 1;
            object[,] ids = new object[n, 1];
            for (int i = 0; i < n; i++) ids[i, 0] = Store.RowId(t, meta, f1 + i);
            Com.SetV2(ws, "A" + f1 + ":A" + l1, ids);
            Com.Set(Com.Range(ws, "L" + f1 + ":L" + l1), "Formula", "=IF(J" + f1 + "=\"\",\"\",IF(M" + f1 + "=\"Pagado\",0,MAX(0,J" + f1 + "-N(K" + f1 + "))))");
            Com.Set(Com.Range(ws, "T" + f1 + ":T" + l1), "Formula", "=IF(OR(O" + f1 + "=\"\",N" + f1 + "=\"Finalizado\",N" + f1 + "=\"Desfavorable\"),\"\",TODAY()-O" + f1 + ")");
            Com.Set(Com.Range(ws, "W" + f1 + ":W" + l1), "Formula", "=IF(J" + f1 + "=\"\",\"\",J" + f1 + "-N(V" + f1 + "))");
            // formatos (solo aspecto: si Excel no los acepta, la pestaña se crea igual)
            try
            {
                foreach (string c in new string[] { "O", "P" }) Com.Set(Com.Range(ws, c + f1 + ":" + c + l1), "NumberFormat", "dd/mm/yyyy");
                foreach (string c in new string[] { "J", "K", "L", "V", "W" }) Com.Set(Com.Range(ws, c + f1 + ":" + c + l1), "NumberFormat", "#,##0.00 \"€\"");
                Com.Set(Com.Range(ws, "T" + f1 + ":T" + l1), "NumberFormat", "0");
            }
            catch (Exception e) { if (IsBusy(e)) throw; App.Log("pestaña nueva: formatos no aplicados (" + e.Message + ")"); }
            try
            {
                Com.Set(Com.Get(ws, "Columns", "A:" + t.LastCol), "ColumnWidth", 16);
                Com.Set(Com.Get(ws, "Columns", "B"), "ColumnWidth", 32); Com.Set(Com.Get(ws, "Columns", "R"), "ColumnWidth", 40);
            }
            catch (Exception e) { if (IsBusy(e)) throw; }
            App.Log("pestaña creada en el Excel: " + t.Sheet);
            return ws;
        }

        static readonly string RangeRx = @"^(.+)!(\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?)$";

        static object ListRange(object ws, Match m) { return Com.Range(Com.Get(Com.Get(Com.Get(ws, "Parent"), "Worksheets"), "Item", m.Groups[1].Value.Trim('\'')), m.Groups[2].Value); }

        public static string[] ReadValidationList(object ws, string col)
        {
            Tramite T = G.T;
            string f;
            try { f = Ps.Str(Com.Get(Com.Get(Com.Range(ws, col + T.FirstRow), "Validation"), "Formula1")); } catch { return new string[0]; }
            if (f.Length == 0) return new string[0];
            f = f.TrimStart('=');
            Match m = PsH.M(f, RangeRx);
            List<string> outp = new List<string>();
            if (m.Success)
            {
                foreach (object c in Com.Cells(ListRange(ws, m))) { string v = Ps.Str(Com.Get(c, "Value2")).Trim(); if (v.Length > 0) outp.Add(v); }
                return outp.ToArray();
            }
            if (PsH.Is(f, @"^\$?[A-Z]+\$?\d+")) return new string[0];
            // «1 PASO, APODERAMIENTO HECHO»: un trozo que empieza por espacio es parte de la opción anterior
            foreach (string piece in f.Split(new char[] { ',', ';' }))
            {
                string v = piece.Trim().Trim('"').Trim();
                if (v.Length == 0) continue;
                if (piece.StartsWith(" ") && outp.Count > 0) outp[outp.Count - 1] = outp[outp.Count - 1] + ", " + v; else outp.Add(v);
            }
            return outp.ToArray();
        }

        static Dictionary<string, double> ReadTariffAt(object ws, string col, int offset)
        {
            Tramite T = G.T;
            Dictionary<string, double> map = new Dictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            string f;
            try { f = Ps.Str(Com.Get(Com.Get(Com.Range(ws, col + T.FirstRow), "Validation"), "Formula1")).TrimStart('='); } catch { return map; }
            Match m = PsH.M(f, RangeRx);
            if (m.Success)
            {
                foreach (object c in Com.Cells(ListRange(ws, m)))
                {
                    string k = Ps.Str(Com.Get(c, "Value2")).Trim();
                    object p = Com.Get(Com.Get(c, "Offset", 0, offset), "Value2");
                    if (k.Length > 0 && p is double) map[k] = (double)p;
                }
            }
            return map;
        }
        public static Dictionary<string, double> ReadTariff(object ws, string col) { return ReadTariffAt(ws, col, 1); }
        public static Dictionary<string, double> ReadTariffNet(object ws, string col) { return ReadTariffAt(ws, col, 2); }

        public static double? ReadTramitePrice(object wb, string key)
        {
            object aj;
            try { aj = Com.Get(Com.Get(wb, "Worksheets"), "Item", "AJUSTES"); } catch { return null; }
            object[,] cells = (object[,])Com.V2(aj, "B1:C200");
            for (int r = 1; r <= 200; r++)
                if (Ps.EqI(Ps.Str(cells[r, 1]).Trim(), key) && cells[r, 2] is double) return (double)cells[r, 2];
            return null;
        }

        // Read-TramiteData: lee la pestaña entera del trámite actual
        public static TramiteData ReadTramiteData(Book h)
        {
            Tramite T = G.T;
            object ws = GetSheet(h, false);
            Dictionary<string, string[]> lists = new Dictionary<string, string[]>();
            foreach (Field f in T.Fields)
                if ((f.List ?? "").StartsWith("excel:", StringComparison.OrdinalIgnoreCase)) { string c = f.List.Substring(6); lists[f.List] = Run<string[]>(delegate { return ReadValidationList(ws, c); }); }
            Dictionary<string, double> tariff = new Dictionary<string, double>();
            if (!string.IsNullOrEmpty(T.TipoCol)) tariff = Run<Dictionary<string, double>>(delegate { return ReadTariff(ws, T.TipoCol); });
            double? defPrice = null;
            if (!string.IsNullOrEmpty(T.PriceKey)) defPrice = Run<double?>(delegate { return ReadTramitePrice(h.Wb, T.PriceKey); });
            string addr = "A" + T.FirstRow + ":" + T.LastCol + T.LastRow;
            object[,] vals = Run<object[,]>(delegate { return (object[,])Com.V2(ws, addr); });
            int n = T.LastRow - T.FirstRow + 1;
            List<ClientRow> rows = new List<ClientRow>();
            int? firstEmpty = null; string nextId = null;
            List<string> payExtra = new List<string>();
            int nameC = Xlsx.ColNum(T.NameCol), idC = Xlsx.ColNum(T.IdCol);
            for (int i = 1; i <= n; i++)
            {
                string name = Ps.Str(vals[i, nameC]).Trim();
                if (name.Length == 0)
                {
                    if (!firstEmpty.HasValue) { firstEmpty = T.FirstRow + i - 1; nextId = Ps.Str(vals[i, idC]); }
                    continue;
                }
                Dictionary<string, object> rec = new Dictionary<string, object>(StringComparer.CurrentCultureIgnoreCase);
                rec["Row"] = T.FirstRow + i - 1;
                foreach (Field f in T.Fields)
                {
                    object v = vals[i, Xlsx.ColNum(f.Col)];
                    if (f.Type == "Date" && v is double) v = DateTime.FromOADate((double)v);
                    rec[f.Key] = v;
                }
                if (!Ps.Blank(Store.Get(rec, "forma"))) { string fo = Ps.Str(rec["forma"]).Trim(); if (!payExtra.Contains(fo)) payExtra.Add(fo); }
                string status = Ps.Str(Store.Get(rec, "status"));
                object pend = vals[i, Xlsx.ColNum(T.PendCol)];
                object days = vals[i, Xlsx.ColNum(T.DaysCol)];
                string color = "#FFFFFF"; if (T.StatusColors.ContainsKey(status)) color = T.StatusColors[status];
                ClientRow r = new ClientRow();
                r.Row = T.FirstRow + i - 1; r.Id = Ps.Str(vals[i, idC]); r.Name = name; r.Doc = Ps.Str(Store.Get(rec, "doc"));
                r.Tipo = rec.ContainsKey("tipo") ? Ps.Str(rec["tipo"]) : Ps.Str(Store.Get(rec, "xF"));
                r.Status = status; r.Color = color; r.Pay = Ps.Str(Store.Get(rec, "paystate"));
                r.Pending = pend is double ? (double)pend : 0;
                r.PendingText = (pend is double && (double)pend > 0) ? Ps.Money(pend) : "";
                r.Days = days is double ? (object)(int)Math.Round((double)days) : "";
                r.Loc = Ps.Str(Store.Get(rec, "loc")); r.Resp = Ps.Str(Store.Get(rec, "resp")); r.Data = rec;
                r.PriceV = vals[i, 10] is double ? (double)vals[i, 10] : 0;
                r.PaidV = vals[i, 11] is double ? (double)vals[i, 11] : 0;
                r.NetV = (!string.IsNullOrEmpty(T.NetCol) && vals[i, Xlsx.ColNum(T.NetCol)] is double) ? (double)vals[i, Xlsx.ColNum(T.NetCol)] : 0;
                r.Fecha = Store.Get(rec, "fsol") is DateTime ? (DateTime?)(DateTime)rec["fsol"] : null;
                r.Sync = true;
                rows.Add(r);
            }
            List<string> all = new List<string>(Store.FormasPago);
            foreach (string p in payExtra) if (!all.Contains(p) && p.Length <= 30) all.Add(p);
            lists["payments"] = all.ToArray();
            lists["nationalities"] = Store.Nacionalidades;
            List<string> col = new List<string>();
            foreach (string c in Config.Cur.Colleagues) if (c.Trim().Length > 0 && !col.Contains(c)) col.Add(c);
            foreach (ClientRow r in rows) { string x = r.Resp.Trim().ToUpper(); if (x.Length > 0 && !col.Contains(x)) col.Add(x); }
            lists["colleagues"] = col.ToArray();
            TramiteData d = new TramiteData();
            d.Lists = lists; d.Tariff = tariff; d.DefaultPrice = defPrice; d.Rows = rows; d.FirstEmpty = firstEmpty; d.NextId = nextId;
            return d;
        }

        // Restore-Formula: vuelve a poner la fórmula automática de esa columna copiándola de OTRA FILA DE CLIENTE
        public static bool RestoreFormula(object ws, string col, int row)
        {
            Tramite T = G.T;
            object fx;
            try { fx = Com.Call(Com.Range(ws, col + T.FirstRow + ":" + col + T.LastRow), "SpecialCells", -4123); }
            catch (Exception e) { if (IsBusy(e)) throw; return false; }
            foreach (object src in Com.Cells(fx))
            {
                int sr = Convert.ToInt32(Com.Get(src, "Row"));
                if (sr == row) continue;
                if (Ps.Str(Com.Get(src, "Formula")).TrimStart('=').ToUpperInvariant().StartsWith("SUM(")) continue;
                if (Com.S(ws, T.NameCol + sr).Trim().Length == 0) continue;
                Com.Set(Com.Range(ws, col + row), "FormulaR1C1", Com.Get(src, "FormulaR1C1")); return true;
            }
            object own = Com.Range(ws, col + row);
            object hf = Com.Get(own, "HasFormula");
            return hf is bool && (bool)hf && !Ps.Str(Com.Get(own, "Formula")).TrimStart('=').ToUpperInvariant().StartsWith("SUM(");
        }

        public static void SetCellValue(object ws, string addr, object value, string type)
        {
            object rng = Com.Range(ws, addr);
            if (value == null || Ps.Str(value) == "") { Com.Call(rng, "ClearContents"); return; }
            if (type == "Date" || type == "Money") Com.Set(rng, "Value2", Convert.ToDouble(value));
            else Com.Set(rng, "Value2", "'" + Ps.Str(value));     // el apóstrofo obliga a guardarlo como texto
        }

        // Save-Client: primero directamente en el archivo; si no se puede con seguridad, con Excel
        public static FileResult SaveClient(Dictionary<string, object> vals, int? row, string expectName, bool isNew, Book handle, bool fileOnly)
        {
            Tramite T = G.T;
            List<ClientOp> ops = ExcelFile.ClientOps(T, vals, G.Tariff);
            if (handle == null)
            {
                FileResult fr = null;
                try { fr = ExcelFile.SaveOne(T, ops, row, expectName, isNew); } catch (Exception e) { App.Log("Excel directo no posible: " + e.Message); }
                if (fr != null) return fr;
                if (fileOnly) return null;
            }
            Book h = handle; bool own = handle == null;
            try
            {
                if (own) h = Run<Book>(delegate { return Open(true); });
                object ws = GetSheet(h, true);
                if (own && Config.Cur.Backup) ExcelFile.BackupExcel();
                if (row.HasValue && row.Value != 0 && isNew)
                {
                    string curName = Com.S(ws, T.NameCol + row.Value).Trim();
                    if (curName.Length > 0 && curName != (expectName ?? "").Trim()) row = null;
                }
                if (!row.HasValue || row.Value == 0)
                {
                    row = null;
                    object[,] col = Run<object[,]>(delegate { return (object[,])Com.V2(ws, T.NameCol + T.FirstRow + ":" + T.NameCol + T.LastRow); });
                    int n = T.LastRow - T.FirstRow + 1;
                    for (int i = 1; i <= n; i++) if (Ps.Str(col[i, 1]).Trim().Length == 0) { row = T.FirstRow + i - 1; break; }
                    if (!row.HasValue) throw new GestemException(Tx.L("errFull"));
                }
                int rr = row.Value;
                Run(delegate
                {
                    foreach (ClientOp op in ops)
                    {
                        // sin fórmula en esa columna (pestaña creada por GESTEM): se escribe el valor calculado
                        if (op.Kind == "formula" && RestoreFormula(ws, op.Col, rr)) continue;
                        SetCellValue(ws, op.Col + rr, op.Value, op.Type);
                    }
                });
                string id = Com.S(ws, T.IdCol + rr);
                if (own) { Book hh = h; Run(delegate { Close(hh, true); }); }
                h = null;
                App.Log("Excel: fila " + rr + " (" + id + ")");
                FileResult res = new FileResult(); res.Id = id; res.Row = rr;
                return res;
            }
            catch
            {
                if (own && h != null) { try { Close(h, false); } catch { } }
                throw;
            }
        }
    }
}

// Datos de GESTEM («GESTEM datos\<trámite>.json» junto al Excel): leer, guardar, bloqueo entre PCs,
// copia diaria y los cálculos de cada fila. Copiado de Read-Store / Write-Store / Resolve-Record / New-RowObject.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Gestem
{
    public class Meta
    {
        public Dictionary<string, string[]> Lists = new Dictionary<string, string[]>();
        public Dictionary<string, double> Tariff = new Dictionary<string, double>();
        public Dictionary<string, double> TariffNet = new Dictionary<string, double>();
        public double? DefaultPrice;
        public string IdPrefix = "";
        public int IdWidth = 3;
    }

    public class DelMark { public string By = "", At = ""; }     // cliente eliminado (marca «borrado»)
    public class HandDelRec { public string Uid = "", Key = "", Name = "", By = "", At = ""; public bool On; }   // eliminado de una pestaña a mano

    public class ClientRec
    {
        public int Row;
        public string Id = "", Updated = "", By = "";
        public string Uid, H;          // id en Firebase y huella de lo último guardado/recibido (null = cliente nuevo)
        public DelMark Del;
        public bool Sync, New;
        public Dictionary<string, object> F = new Dictionary<string, object>();
    }

    public class StoreData
    {
        public Meta Meta = new Meta();
        public List<ClientRec> Clients = new List<ClientRec>();
        public List<string> Adopted = new List<string>();
        public List<HandDelRec> HandDel = new List<HandDelRec>();
    }

    // una fila de la lista de clientes (las propiedades son las que usa la tabla de la ventana)
    public class ClientRow
    {
        public int Row { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string Doc { get; set; }
        public string Tipo { get; set; }
        public string Status { get; set; }
        public string Color { get; set; }
        public string Pay { get; set; }
        public double Pending { get; set; }
        public string PendingText { get; set; }
        public object Days { get; set; }
        public string Loc { get; set; }
        public string Resp { get; set; }
        public Dictionary<string, object> Data { get; set; }
        public double PriceV { get; set; }
        public double PaidV { get; set; }
        public double NetV { get; set; }
        public DateTime? Fecha { get; set; }
        public bool Sync { get; set; }
        public string Updated { get; set; }
        public string By { get; set; }
        public int DocsN { get; set; }
        public string DocsVis { get; set; }
    }

    public class TramiteData
    {
        public Dictionary<string, string[]> Lists = new Dictionary<string, string[]>();
        public Dictionary<string, double> Tariff = new Dictionary<string, double>();
        public double? DefaultPrice;
        public List<ClientRow> Rows = new List<ClientRow>();
        public int? FirstEmpty;
        public string NextId = "";
        public StoreData Store;
    }

    public static class Store
    {
        public static string Dir()
        {
            string x = Config.Cur.ExcelPath ?? "";
            if (x.Length == 0) return null;
            return Path.Combine(Path.GetDirectoryName(x), "GESTEM datos");
        }
        public static string PathOf(Tramite t) { string d = Dir(); return d == null ? null : Path.Combine(d, t.Id + ".json"); }

        // Get-FileStamp: si otro PC guarda, se vuelve a leer
        public static long FileStamp()
        {
            string d = Dir();
            try
            {
                if (d != null && Directory.Exists(d))
                {
                    long max = 0;
                    foreach (string f in Directory.GetFiles(d, "*.json"))
                    {
                        long t = File.GetLastWriteTimeUtc(f).Ticks;
                        if (t > max) max = t;
                    }
                    return max;
                }
            }
            catch { }
            return 0;
        }

        // Enter-StoreLock / Exit-StoreLock (bloqueo.lock en la carpeta de datos)
        public static string EnterLock()
        {
            string d = Dir();
            if (!Directory.Exists(d)) Directory.CreateDirectory(d);
            string lk = Path.Combine(d, "bloqueo.lock");
            for (int i = 0; i < 50; i++)
            {
                try
                {
                    using (FileStream fs = new FileStream(lk, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    {
                        byte[] b = Encoding.UTF8.GetBytes(Environment.MachineName + " " + App.User + " " + DateTime.Now.ToString("s", Ps.Inv));
                        fs.Write(b, 0, b.Length);
                    }
                    return lk;
                }
                catch
                {
                    // un bloqueo de más de 1 minuto es de un programa que se cerró mal
                    try { if ((DateTime.Now - File.GetLastWriteTime(lk)).TotalSeconds > 60) File.Delete(lk); } catch { }
                    System.Threading.Thread.Sleep(200);
                }
            }
            throw new GestemException(Tx.L("errStoreBusy"));
        }
        public static void ExitLock(string lk) { if (lk != null) { try { File.Delete(lk); } catch { } } }

        public static object ToStoreValue(Field f, object v)
        {
            if (Ps.Blank(v)) return null;
            if (f.Type == "Date") { if (v is DateTime) return ((DateTime)v).ToString("yyyy-MM-dd", Ps.Inv); return Ps.Str(v); }
            if (f.Type == "Money") { if (v is double || v is int || v is decimal) return Convert.ToDouble(v); double? d = Ps.ParseMoney(Ps.Str(v)); return d.HasValue ? (object)d.Value : null; }
            if (v is double) return v;
            return Ps.Str(v);
        }
        public static object FromStoreValue(Field f, object v)
        {
            if (Ps.Blank(v)) return null;
            if (f.Type == "Date")
            {
                if (v is DateTime) return v;
                DateTime d;
                if (DateTime.TryParseExact(Ps.Str(v), "yyyy-MM-dd", Ps.Inv, System.Globalization.DateTimeStyles.None, out d)) return d;
                return null;
            }
            if (f.Type == "Money") { double? n = Ps.Num(v); return n.HasValue ? (object)n.Value : null; }   // texto («CNTA …») = sin importe
            return v;
        }

        public static StoreData Read(Tramite t)
        {
            string p = PathOf(t);
            if (p == null || !File.Exists(p)) return null;
            object j = Json.Parse(File.ReadAllText(p, Encoding.UTF8));
            object jm = Json.Get(j, "meta");
            StoreData s = new StoreData();
            Meta m = s.Meta;
            m.IdPrefix = Ps.Str(Json.Get(jm, "idPrefix"));
            object iw = Json.Get(jm, "idWidth");
            if (iw != null && Ps.Num(iw).GetValueOrDefault() != 0) m.IdWidth = Convert.ToInt32(Ps.Num(iw).Value);
            Dictionary<string, object> lists = Json.Obj(Json.Get(jm, "lists"));
            if (lists != null) foreach (KeyValuePair<string, object> kv in lists) m.Lists[kv.Key] = Ps.StrList(kv.Value).ToArray();
            Dictionary<string, object> tar = Json.Obj(Json.Get(jm, "tariff"));
            if (tar != null) foreach (KeyValuePair<string, object> kv in tar) m.Tariff[kv.Key] = Convert.ToDouble(kv.Value);
            Dictionary<string, object> tn = Json.Obj(Json.Get(jm, "tariffNet"));
            if (tn != null) foreach (KeyValuePair<string, object> kv in tn) m.TariffNet[kv.Key] = Convert.ToDouble(kv.Value);
            object dp = Json.Get(jm, "defaultPrice");
            if (dp != null) m.DefaultPrice = Convert.ToDouble(dp);
            HashSet<string> seen = new HashSet<string>();
            foreach (object c in Json.Arr(Json.Get(j, "clients")))
            {
                if (c == null) continue;
                ClientRec r = new ClientRec();
                object jf = Json.Get(c, "f");
                foreach (Field f in t.Fields) r.F[f.Key] = FromStoreValue(f, Json.Get(jf, f.Key));
                r.Row = Convert.ToInt32(Json.Get(c, "row") ?? 0L);
                // id único para Firebase; los clientes de antes de Firebase lo sacan de su nº, así todos los PC dan el mismo
                r.Uid = Ps.Str(Json.Get(c, "uid"));
                object del = Json.Get(c, "borrado");
                if (r.Row < Hand.RowBase)
                {
                    string id0 = Ps.Str(Json.Get(c, "id"));
                    if (r.Uid.Length == 0) r.Uid = PsH.Rep(id0.Length > 0 ? t.Id + "-" + id0 : t.Id + "-r" + r.Row, @"[^\w-]", "_");
                    if (!seen.Add(r.Uid)) r.Uid = Guid.NewGuid().ToString("N");
                    List<object> raw = new List<object>(); foreach (Field f in FbFields(t)) raw.Add(Json.Get(jf, f.Key));
                    r.H = FbHash(FbKey(id0, Truthy(del)), raw);
                }
                if (Truthy(del)) { r.Del = new DelMark(); r.Del.By = Ps.Str(Json.Get(del, "by")); r.Del.At = Ps.Str(Json.Get(del, "at")); }
                r.Id = Ps.Str(Json.Get(c, "id"));
                r.Sync = Truthy(Json.Get(c, "sync"));
                r.New = Truthy(Json.Get(c, "nuevo"));
                r.Updated = Ps.Str(Json.Get(c, "updated"));
                r.By = Ps.Str(Json.Get(c, "by"));
                s.Clients.Add(r);
            }
            foreach (object a in Json.Arr(Json.Get(j, "adoptados"))) if (Ps.Str(a).Trim().Length > 0) s.Adopted.Add(Ps.Str(a));
            // clientes de pestañas hechas a mano eliminados en GESTEM (por nombre); On = false si se deshizo
            foreach (object e in Json.Arr(Json.Get(j, "borradosMano")))
            {
                if (e == null || !Truthy(Json.Get(e, "uid"))) continue;
                HandDelRec x = new HandDelRec(); x.Uid = Ps.Str(Json.Get(e, "uid")); x.Key = Ps.Str(Json.Get(e, "key")); x.Name = Ps.Str(Json.Get(e, "name"));
                x.On = Truthy(Json.Get(e, "on")); x.By = Ps.Str(Json.Get(e, "by")); x.At = Ps.Str(Json.Get(e, "at"));
                s.HandDel.Add(x);
            }
            return s;
        }
        static bool Truthy(object v) { if (v == null) return false; if (v is bool) return (bool)v; if (v is string) return ((string)v).Length > 0; if (Ps.IsNum(v)) return Convert.ToDouble(v) != 0; return true; }

        public static void Write(Tramite t, StoreData store)
        {
            string p = PathOf(t);
            List<ClientRec> sorted = new List<ClientRec>(store.Clients);
            StableSortByRow(sorted);
            List<object> list = new List<object>();
            Dictionary<int, ClientRec> prev = null; FbState st = null;
            foreach (ClientRec c in sorted)
            {
                Dictionary<string, object> o = new Dictionary<string, object>();
                foreach (Field f in t.Fields) { object v; c.F.TryGetValue(f.Key, out v); o[f.Key] = ToStoreValue(f, v); }
                if (c.Row < Hand.RowBase)
                {
                    if (string.IsNullOrEmpty(c.Uid))
                    {
                        // cliente nuevo o vuelto a importar del Excel: conserva el id si es el mismo cliente en la misma fila
                        if (prev == null)
                        {
                            prev = new Dictionary<int, ClientRec>();
                            try { StoreData ps = Read(t); if (ps != null) foreach (ClientRec x0 in ps.Clients) prev[x0.Row] = x0; } catch { }
                        }
                        ClientRec x; prev.TryGetValue(c.Row, out x);
                        if (x != null && !string.IsNullOrEmpty(x.Uid) && Ps.Str(Get(x.F, "name")) == Ps.Str(Get(c.F, "name"))) { c.Uid = x.Uid; c.H = x.H; }
                        else c.Uid = Guid.NewGuid().ToString("N");
                    }
                    // cambiado en este PC: queda «por enviar» a Firebase (lo que llega de Firebase no se reenvía)
                    List<object> vals = new List<object>(); foreach (Field f in FbFields(t)) vals.Add(o[f.Key]);
                    string h = FbHash(FbKey(c.Id, c.Del != null), vals);
                    if (h != c.H && !FbState.Applying)
                    {
                        if (st == null) st = FbState.Load();
                        PendItem pi = new PendItem(); pi.T = t.Id; pi.H = h; pi.N = st.Pend.ContainsKey(c.Uid) && ((PendItem)st.Pend[c.Uid]).N;
                        st.Pend[c.Uid] = pi;
                    }
                    c.H = h;
                }
                Dictionary<string, object> rec = new Dictionary<string, object>();
                rec["row"] = (long)c.Row; rec["id"] = c.Id; rec["uid"] = c.Uid; rec["sync"] = c.Sync; rec["nuevo"] = c.New;
                rec["updated"] = c.Updated; rec["by"] = c.By; rec["f"] = o;
                if (c.Del != null) { Dictionary<string, object> b = new Dictionary<string, object>(); b["by"] = c.Del.By; b["at"] = c.Del.At; rec["borrado"] = b; }
                list.Add(rec);
            }
            Meta m = store.Meta;
            Dictionary<string, object> meta = new Dictionary<string, object>();
            meta["idPrefix"] = m.IdPrefix; meta["idWidth"] = (long)m.IdWidth;
            meta["defaultPrice"] = m.DefaultPrice.HasValue ? (object)m.DefaultPrice.Value : null;
            meta["tariff"] = AsPsH(m.Tariff); meta["tariffNet"] = AsPsH(m.TariffNet); meta["lists"] = AsPsH(m.Lists);    // @{} en PowerShell: su orden al guardar
            List<string> adopted = new List<string>();
            foreach (string a in store.Adopted) if ((a ?? "").Trim().Length > 0) adopted.Add(a);
            Dictionary<string, object> obj = new Dictionary<string, object>();
            obj["app"] = App.Name; obj["formato"] = 1L; obj["tramite"] = t.Id; obj["hoja"] = t.Sheet;
            obj["guardado"] = DateTime.Now.ToString("s", Ps.Inv);
            obj["meta"] = meta; obj["clients"] = list; obj["adoptados"] = adopted;
            List<object> hd = new List<object>();
            foreach (HandDelRec e in store.HandDel)
            {
                if (e == null) continue;
                Dictionary<string, object> x = new Dictionary<string, object>();
                x["uid"] = e.Uid; x["key"] = e.Key; x["name"] = e.Name; x["on"] = e.On; x["by"] = e.By; x["at"] = e.At; hd.Add(x);
            }
            obj["borradosMano"] = hd;
            string json = Json.Write(obj);
            string tmp = p + ".tmp";
            File.WriteAllText(tmp, json, new UTF8Encoding(true));
            if (File.Exists(p))
            {
                // copia diaria del archivo de datos antes de cambiarlo
                try
                {
                    string bk = Path.Combine(App.BackupDir, "datos_" + t.Id + "_" + DateTime.Now.ToString("yyyy-MM-dd", Ps.Inv) + ".json");
                    if (!File.Exists(bk)) File.Copy(p, bk, true);
                }
                catch { }
                try { File.Replace(tmp, p, null); }
                catch { File.Copy(tmp, p, true); try { File.Delete(tmp); } catch { } }
            }
            else File.Move(tmp, p);
            if (st != null) st.Save();
        }

        // Get-FbFields / Get-FbKey / Get-FbHash: lo que se sube a Firebase (sin contraseñas) y su huella
        public static List<Field> FbFields(Tramite t) { List<Field> r = new List<Field>(); foreach (Field f in t.Fields) if (!f.Local) r.Add(f); return r; }
        public static string FbKey(string id, bool del) { return del ? (id ?? "") + (char)30 + "BORRADO" : (id ?? ""); }
        public static string FbHash(string id, List<object> vals)
        {
            // -join de PowerShell: números con la cultura del PC (es-ES: «12,5»)
            List<string> p = new List<string>();
            foreach (object v in vals) p.Add(v == null ? "" : v is string ? (string)v : v is bool ? ((bool)v ? "True" : "False") : Convert.ToString(v, System.Globalization.CultureInfo.CurrentCulture));
            string s = id + (char)31 + string.Join(((char)31).ToString(), p.ToArray());
            using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create())
                return Convert.ToBase64String(md5.ComputeHash(Encoding.UTF8.GetBytes(s)));
        }

        static System.Collections.Hashtable AsPsH<V>(Dictionary<string, V> d) { System.Collections.Hashtable h = PsH.New(); foreach (KeyValuePair<string, V> kv in d) h[kv.Key] = kv.Value; return h; }

        public static void StableSortByRow(List<ClientRec> l)
        {
            // Sort-Object de PowerShell es estable; List.Sort no: se ordena con el índice como desempate
            List<KeyValuePair<int, ClientRec>> tmp = new List<KeyValuePair<int, ClientRec>>();
            for (int i = 0; i < l.Count; i++) tmp.Add(new KeyValuePair<int, ClientRec>(i, l[i]));
            tmp.Sort(delegate(KeyValuePair<int, ClientRec> a, KeyValuePair<int, ClientRec> b)
            {
                int c = a.Value.Row.CompareTo(b.Value.Row);
                return c != 0 ? c : a.Key.CompareTo(b.Key);
            });
            l.Clear();
            foreach (KeyValuePair<int, ClientRec> kv in tmp) l.Add(kv.Value);
        }

        // --- valores que en el Excel calculan sus fórmulas ---
        public static Dictionary<string, object> Resolve(Tramite t, Dictionary<string, object> FV, Meta meta)
        {
            UpdateReqDeadline(t, FV);
            Field natF = t.FieldByRole("nationality");
            if (natF != null && Ps.Trim(Get(FV, natF.Key)).Length == 0 && Ps.Trim(Get(FV, "name")).Length > 0) FV[natF.Key] = t.DefaultNationality;
            string tipo = Ps.Str(Get(FV, "tipo"));
            foreach (string k in new string[] { "price", "paid", "gastos" })
                if (FV.ContainsKey(k)) { double? n = Ps.Num(FV[k]); FV[k] = n.HasValue ? (object)n.Value : null; }
            if (!string.IsNullOrEmpty(t.TipoCol) && Get(FV, "price") == null && tipo.Length > 0 && meta.Tariff.ContainsKey(tipo)) FV["price"] = meta.Tariff[tipo];
            object price = Get(FV, "price");
            string ps = Ps.Str(Get(FV, "paystate"));
            if (price == null) { if (ps != "Parcial") FV["paid"] = null; }
            else if (ps == "Pagado") FV["paid"] = Convert.ToDouble(price);
            else if (ps == "Pendiente") FV["paid"] = 0.0;
            else if (ps != "Parcial") FV["paid"] = null;
            return FV;
        }

        public static object Get(Dictionary<string, object> d, string k) { object v; return d != null && d.TryGetValue(k, out v) ? v : null; }

        // --- requerimientos: Get-ReqDeadline / Get-WorkDaysUntil / Update-ReqDeadline / Get-ReqAlerts ---
        // 10 días hábiles (lunes a viernes) desde el día siguiente a la apertura, o a la notificación si no hay apertura
        // shortcut: no cuenta festivos (igual que PowerShell); añadir una lista de festivos si hace falta
        public static DateTime? ReqDeadline(object apertura, object notif)
        {
            object o = apertura is DateTime ? apertura : notif is DateTime ? notif : null;
            if (o == null) return null;
            DateTime d = ((DateTime)o).Date; int n = 0;
            while (n < 10) { d = d.AddDays(1); if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) n++; }
            return d;
        }
        public static int WorkDaysUntil(DateTime from, DateTime to)
        {
            int n = 0; DateTime d = from.Date;
            while (d < to.Date) { d = d.AddDays(1); if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) n++; }
            return n;
        }
        public static void UpdateReqDeadline(Tramite t, Dictionary<string, object> FV)
        {
            if (t.HasField("flim")) { DateTime? d = ReqDeadline(Get(FV, "xG"), Get(FV, "xF")); FV["flim"] = d.HasValue ? (object)d.Value : null; }
        }
        // requerimientos abiertos (sin respuesta ni cerrados) que vencen en 3 días hábiles o menos, y cuántos ya vencieron
        public static void ReqAlerts(IEnumerable<ClientRow> rows, DateTime today, out int due, out int late)
        {
            due = 0; late = 0;
            foreach (ClientRow r in rows)
            {
                if (r == null) continue;
                Dictionary<string, object> F = r.Data;
                string st = Ps.Str(Get(F, "xI"));
                if (st == "Respondido" || st == "Cerrado" || Get(F, "xH") is DateTime) continue;
                DateTime? dl = ReqDeadline(Get(F, "xG"), Get(F, "xF"));
                if (!dl.HasValue) continue;
                if (dl.Value < today.Date) late++; else if (WorkDaysUntil(today, dl.Value) <= 3) due++;
            }
        }
        // columna «Tipo» de la lista: el tipo, o la fecha de notificación (requerimientos)
        public static string TipoText(Dictionary<string, object> FV)
        {
            if (FV.ContainsKey("tipo")) return Ps.Str(Get(FV, "tipo"));
            object x = Get(FV, "xF");
            return x is DateTime ? ((DateTime)x).ToString("dd/MM/yyyy", Ps.Inv) : Ps.Str(x);
        }

        public static ClientRow RowObject(Tramite t, ClientRec c, Meta meta)
        {
            Dictionary<string, object> FV = c.F;
            Dictionary<string, object> rec = new Dictionary<string, object>();
            rec["Row"] = c.Row;
            foreach (KeyValuePair<string, object> kv in FV) rec[kv.Key] = kv.Value;
            double? price = Ps.Num(Get(FV, "price")); double? paid = Ps.Num(Get(FV, "paid"));
            string ps = Ps.Str(Get(FV, "paystate")); string status = Ps.Str(Get(FV, "status"));
            double pend = 0.0;
            if (price.HasValue && ps != "Pagado") { double pv = paid.HasValue ? paid.Value : 0.0; pend = Math.Max(0.0, price.Value - pv); }
            object days = "";
            bool closed = status == "Finalizado" || (string.IsNullOrEmpty(t.TipoCol) && status == "Desfavorable");
            object fsol = Get(FV, "fsol");
            if (fsol is DateTime && !closed) days = (int)Math.Round((DateTime.Today - ((DateTime)fsol).Date).TotalDays);
            double net = 0.0;
            if (price.HasValue)
            {
                if (t.NetCol == "W") { double g = 0.0; double? gv = Ps.Num(Get(FV, "gastos")); if (gv.HasValue) g = gv.Value; net = price.Value - g; }
                else if (!string.IsNullOrEmpty(t.TipoCol) && meta.TariffNet.ContainsKey(Ps.Str(Get(FV, "tipo")))) net = meta.TariffNet[Ps.Str(Get(FV, "tipo"))];
            }
            string color = "#FFFFFF";
            if (t.StatusColors.ContainsKey(status)) color = t.StatusColors[status];
            ClientRow r = new ClientRow();
            r.Row = c.Row; r.Id = c.Id; r.Name = Ps.Str(Get(FV, "name")); r.Doc = Ps.Str(Get(FV, "doc"));
            r.Tipo = TipoText(FV);
            r.Status = status; r.Color = color; r.Pay = ps;
            r.Pending = pend; r.PendingText = pend > 0 ? Ps.Money(pend) : "";
            r.Days = days; r.Loc = Ps.Str(Get(FV, "loc")); r.Resp = Ps.Str(Get(FV, "resp")); r.Data = rec;
            r.PriceV = price.HasValue ? price.Value : 0; r.PaidV = paid.HasValue ? paid.Value : 0; r.NetV = net;
            r.Fecha = fsol is DateTime ? (DateTime?)(DateTime)fsol : null;
            r.Sync = c.Sync; r.Updated = c.Updated; r.By = c.By;
            return r;
        }

        public static string RowId(Tramite t, Meta meta, int row)
        {
            if (string.IsNullOrEmpty(meta.IdPrefix)) return "";
            return meta.IdPrefix + (row - t.FirstRow + 1).ToString(Ps.Inv).PadLeft(meta.IdWidth, '0');
        }

        public static readonly string[] Nacionalidades = Ps.StrList(Json.Get(App.ResJson("lists.json"), "nationalities")).ToArray();
        public static readonly string[] FormasPago = Ps.StrList(Json.Get(App.ResJson("lists.json"), "payments")).ToArray();

        public static TramiteData BuildData(Tramite t, StoreData store)
        {
            Meta meta = store.Meta;
            List<ClientRow> rows = new List<ClientRow>();
            HashSet<int> used = new HashSet<int>();
            List<string> payExtra = new List<string>();      // HashSet de PowerShell: mismo orden de inserción
            List<ClientRec> sorted = new List<ClientRec>(store.Clients);
            StableSortByRow(sorted);
            foreach (ClientRec c in sorted)
            {
                used.Add(c.Row);
                if (c.Del != null) continue;    // eliminado: no se ve, pero su nº no se vuelve a usar
                rows.Add(RowObject(t, c, meta));
                string fo = Ps.Trim(Get(c.F, "forma"));
                if (fo.Length > 0 && !payExtra.Contains(fo)) payExtra.Add(fo);
            }
            int? firstEmpty = null;
            for (int r = t.FirstRow; r <= t.LastRow; r++) if (!used.Contains(r)) { firstEmpty = r; break; }
            Dictionary<string, string[]> lists = new Dictionary<string, string[]>();
            foreach (KeyValuePair<string, string[]> kv in meta.Lists) lists[kv.Key] = kv.Value;
            // listas guardadas antes de 1.5.14: «1 PASO» + «APODERAMIENTO HECHO» se unen si los clientes tienen «1 PASO, APODERAMIENTO HECHO»
            foreach (Field f in t.Fields)
            {
                if (!(f.List ?? "").StartsWith("excel:") || !lists.ContainsKey(f.List)) continue;
                HashSet<string> vals = new HashSet<string>();
                foreach (ClientRec c in store.Clients) { string v = Ps.Trim(Get(c.F, f.Key)); if (v.Length > 0) vals.Add(v); }
                string[] src = lists[f.List]; List<string> fixedL = new List<string>(); bool changed = false;
                for (int i = 0; i < src.Length; i++)
                {
                    if (i + 1 < src.Length && !vals.Contains(src[i]) && vals.Contains(src[i] + ", " + src[i + 1])) { fixedL.Add(src[i] + ", " + src[i + 1]); i++; changed = true; }
                    else fixedL.Add(src[i]);
                }
                if (changed) lists[f.List] = fixedL.ToArray();
            }
            List<string> all = new List<string>(FormasPago);
            foreach (string p in payExtra) if (!all.Contains(p) && p.Length <= 30) all.Add(p);
            lists["payments"] = all.ToArray();
            lists["nationalities"] = Nacionalidades;
            List<string> col = new List<string>();
            foreach (string c in Config.Cur.Colleagues) if (c.Trim().Length > 0 && !col.Contains(c)) col.Add(c);
            foreach (ClientRow r in rows) { string x = Ps.Upper(r.Resp.Trim()); if (x.Length > 0 && !col.Contains(x)) col.Add(x); }
            lists["colleagues"] = col.ToArray();
            TramiteData d = new TramiteData();
            d.Lists = lists; d.Tariff = meta.Tariff; d.DefaultPrice = meta.DefaultPrice; d.Rows = rows; d.FirstEmpty = firstEmpty;
            d.NextId = firstEmpty.HasValue ? RowId(t, meta, firstEmpty.Value) : ""; d.Store = store;
            return d;
        }
    }

    public class GestemException : Exception { public GestemException(string m) : base(m) { } }
}

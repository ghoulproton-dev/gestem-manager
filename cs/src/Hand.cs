// Pestañas «hechas a mano» del Excel (AUTÓNOMOS, TASA…): se leen y sus clientes salen en GESTEM.
// Copiado de Get-HandRole / Get-HandLayout / ConvertFrom-HandRow / ConvertFrom-HandSheet / Sync-HandmadeTabs.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Gestem
{
    public class HandLayout { public int HeaderRow; public Hashtable Map = PsH.New(); public int NameCol = 1; public bool PresDate; }

    public static class Hand
    {
        public const int RowBase = 100000;          // fila = 100000 + fila del Excel (solo lectura)
        public static Hashtable HandTabs = PsH.New();   // trámite -> pestaña hecha a mano que se está leyendo
        public static long HandStamp = -1;

        public static Dictionary<string, string[]> SheetAliases = LoadAliases();
        static Dictionary<string, string[]> LoadAliases()
        {
            Dictionary<string, string[]> r = new Dictionary<string, string[]>();
            Dictionary<string, object> d = Json.Obj(Json.Get(App.ResJson("lists.json"), "sheetAliases"));
            if (d != null) foreach (KeyValuePair<string, object> kv in d) r[kv.Key] = Ps.StrList(kv.Value).ToArray();
            return r;
        }
        public static readonly string[] GenStatusList = Ps.StrList(Json.Get(App.ResJson("lists.json"), "genStatus")).ToArray();

        public static string HandKey(string s) { return PsH.Rep(PsH.Rep(Xlsx.SheetKey(s), "[^A-Z0-9 ]", " "), @"\s+", " ").Trim(); }

        public static string DefaultIdPrefix(Tramite t)
        {
            string k = PsH.Rep(Xlsx.SheetKey(t.Sheet), "[^A-Z]", "");
            if (k.Length == 0) return "";
            return k.Substring(0, Math.Min(3, k.Length)) + "-";
        }

        public static StoreData NewEmptyStore(Tramite t, double? price)
        {
            StoreData s = new StoreData();
            s.Meta.Lists["excel:N"] = GenStatusList;
            s.Meta.Lists["excel:M"] = new string[] { "Pagado", "Parcial", "Pendiente" };
            s.Meta.DefaultPrice = price;
            s.Meta.IdPrefix = DefaultIdPrefix(t);
            s.Meta.IdWidth = 3;
            return s;
        }

        static bool RestOk(string k, string key)
        {
            if (!k.StartsWith(key)) return false;
            string rest = k.Substring(key.Length);
            return rest == "S" || rest == "ES" || rest.StartsWith(" ");
        }

        // Get-TabCandidates
        public static string[] TabCandidates(string[] names, Tramite t)
        {
            List<string> res = new List<string>();
            Hashtable keys = PsH.New(); foreach (string n in names) keys[n] = Xlsx.SheetKey(n);
            List<string> wanted = new List<string>(); wanted.Add(t.Sheet);
            if (SheetAliases.ContainsKey(t.Id)) wanted.AddRange(SheetAliases[t.Id]);
            foreach (string w in wanted)
            {
                string k = Xlsx.SheetKey(w);
                foreach (string n in names) if (string.Equals((string)keys[n], k, StringComparison.CurrentCultureIgnoreCase) && !res.Contains(n)) res.Add(n);
            }
            string key = Xlsx.SheetKey(t.Sheet);
            foreach (string n in names) { string k = (string)keys[n]; if (RestOk(k, key) && !res.Contains(n)) res.Add(n); }
            return res.ToArray();
        }

        // Test-CellsGestem
        public static bool CellsGestem(Hashtable cells, Tramite t)
        {
            Hashtable r = cells[t.FirstRow - 1] as Hashtable;
            if (r == null || r.Count == 0) return false;
            string h = Xlsx.SheetKey(Ps.Str(r[Xlsx.ColNum(t.NameCol)]));
            return h.StartsWith("CLIENTE") || h.StartsWith("NOMBRE");
        }

        public static string Role(string k, Tramite t)
        {
            if (PsH.Is(k, "^(NOMBRE|CLIENTE)")) return "name";
            if (PsH.Is(k, @"\b(NIE|DNI|PASAPORTE|DOCUMENTO)\b")) return "doc";
            if (PsH.Is(k, @"TELEFONO|MOVIL|\bTLF\b")) return "tel";
            if (PsH.Is(k, "NACIONALIDAD")) return "nat";
            if (PsH.Is(k, "^(NO PAGADO|PENDIENTE|PENDIETE)")) return "pend";
            if (PsH.Is(k, "^(PAGADO|HA PAGADO|COBRADO|A CUENTA)$")) return "paid";
            if (PsH.Is(k, "^PAGOS?$|FORMA DE PAGO")) return "paytext";
            if (PsH.Is(k, "PRECIO|IMPORTE|HONORARIOS|^TOTAL")) return "price";
            if (PsH.Is(k, "OBSERVACION|NOTAS|COMENTARIO")) return "obs";
            if (PsH.Is(k, "ESTADO|SITUACION")) return "status";
            if (PsH.Is(k, "RESOLUCION")) return "fres";
            if (PsH.Is(k, "FECHA|PRESENTACION|SOLICITUD")) return "fsol";
            // campos propios del trámite (Contraseña, Tipo de gestión…), también por sus opciones (BACHILLERATO -> Título)
            List<string> kw = new List<string>(); foreach (string w in k.Split(' ')) if (w.Length >= 5) kw.Add(w);
            foreach (Field f in t.Fields)
            {
                if (!f.Key.StartsWith("x", StringComparison.OrdinalIgnoreCase)) continue;
                string es; f.Label.TryGetValue("es", out es);
                string lb = HandKey(es);
                if (string.Equals(k, lb, StringComparison.CurrentCultureIgnoreCase) || lb.Contains(k) || k.Contains(lb)) return f.Key;
                foreach (string w in kw) if (PsH.ContainsI(lb.Split(' '), w)) return f.Key;
                if ((f.List ?? "").StartsWith("fixed:", StringComparison.OrdinalIgnoreCase))
                    foreach (string o in f.List.Substring(6).Split('|')) if (string.Equals(HandKey(o), k, StringComparison.CurrentCultureIgnoreCase)) return f.Key;
            }
            return null;
        }

        static List<int> SortedKeys(Hashtable h) { List<int> l = new List<int>(); foreach (object k in h.Keys) l.Add((int)k); l.Sort(); return l; }

        public static HandLayout Layout(Hashtable cells, Tramite t)
        {
            HandLayout lay = new HandLayout();
            int hdr = 0;
            for (int r = 1; r <= 5 && hdr == 0; r++)
            {
                if (!cells.ContainsKey(r)) continue;
                Hashtable row = (Hashtable)cells[r];
                foreach (object c in row.Keys) if (PsH.Is(HandKey(Ps.Str(row[c])), "^(NOMBRE|CLIENTE)")) { hdr = r; break; }
            }
            Hashtable map = lay.Map; bool pres = false;
            if (hdr != 0)
            {
                Hashtable hrow = (Hashtable)cells[hdr];
                foreach (int c in SortedKeys(hrow))
                {
                    string k = HandKey(Ps.Str(hrow[c]));
                    if (k.Length == 0) continue;
                    string role = Role(k, t);
                    if (role != null)
                    {
                        List<string> vals = new List<string>(); foreach (object v in map.Values) vals.Add(Ps.Str(v));
                        if (PsH.ContainsI(vals, role) && !PsH.ContainsI(new string[] { "obs", "paytext", "status" }, role)) role = null;   // la primera columna gana
                    }
                    map[c] = role;
                    if (role == "fsol" && PsH.Is(k, "PRESENTACION")) pres = true;
                    if (role == null) map[c] = "#" + Ps.Str(hrow[c]).Trim();     // cabecera desconocida: va a observaciones con su nombre
                }
            }
            if (hdr == 0 && t.HandCols != null) foreach (KeyValuePair<int, string> e in t.HandCols) map[e.Key] = e.Value;
            int nameCol = 1;
            foreach (DictionaryEntry e in map) if (Ps.Str(e.Value) == "name") { nameCol = (int)e.Key; break; }
            map[nameCol] = "name";
            lay.HeaderRow = hdr; lay.NameCol = nameCol; lay.PresDate = pres;
            return lay;
        }

        public static string Text(object v)
        {
            if (v is double) { double d = (double)v; return d == Math.Floor(d) ? ((long)d).ToString() : d.ToString("0.##", CultureInfo.InvariantCulture); }
            return PsH.Rep(Ps.Str(v), @"\s+", " ").Trim();
        }

        public static object Date(object v)
        {
            if (v is double) { double d = (double)v; if (d >= 20000 && d <= 80000) return DateTime.FromOADate(d).Date; return null; }
            Match m = PsH.M(Ps.Str(v), @"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})");
            if (m.Success) { try { return new DateTime(int.Parse(m.Groups[3].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[1].Value)); } catch { } }
            return null;
        }

        public static double? Number(object v)
        {
            if (v is double) return (double)v;
            Match m = PsH.M(Ps.Str(v), @"(\d+(?:[.,]\d{1,2})?)");
            if (m.Success) return Ps.ParseMoney(m.Groups[1].Value);
            return null;
        }

        public static string Status(List<string> texts, bool hasPresDate)
        {
            string all = HandKey(string.Join(" ", texts.ToArray()));
            if (all.Length == 0 && !hasPresDate) return null;
            if (PsH.Is(all, "FINALIZ|ENTREGAD|TERMINAD")) return "Finalizado";
            if (PsH.Is(all, "DESFAVORABLE|DENEGAD")) return "Desfavorable";
            if (PsH.Is(all, "FAVORABLE|CONCEDID|APROBAD")) return "Favorable";
            if (PsH.Is(all, "REQUERIM")) return "Requerimiento";
            if (PsH.Is(all, "PRESENTAD")) return "Presentado";
            if (PsH.Is(all, "TRAMIT|EN GESTION")) return "En Trámite";
            if (hasPresDate) return "Presentado";
            if (PsH.Is(all, "ESPERA|PREPARAD|DOCUMENTAC|FALTA|PENDIENTE PARA")) return "Documentación";
            return null;
        }

        public static Dictionary<string, object> FromRow(Hashtable row, HandLayout lay, Tramite t)
        {
            Dictionary<string, object> FV = new Dictionary<string, object>();
            foreach (Field f in t.Fields) FV[f.Key] = null;
            List<string> notes = new List<string>(), stTexts = new List<string>(), payTexts = new List<string>();
            double? price = null, paid = null, pend = null; bool presDate = false;
            foreach (int c in SortedKeys(row))
            {
                object v = row[c];
                string role = lay.Map.ContainsKey(c) ? (string)lay.Map[c] : null;
                if (role == "name") continue;
                string txt = Text(v);
                if (txt.Length == 0) continue;
                string rs = role ?? "";
                if (t.Id == "tasas" && v is double && rs != "price" && rs != "paid" && rs != "pend")
                {
                    object dd = Date(v);
                    if (dd != null && rs != "fsol" && rs != "fres") txt = ((DateTime)dd).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                bool matched = false;
                // switch -Wildcard de PowerShell: se ejecutan TODAS las ramas que coinciden
                if (Ps.EqI(rs, "tipoOrDate"))
                {
                    matched = true;
                    if (PsH.Is(HandKey(txt), "^(DT|EX|LARGA|HIJO|TASA|MODELO|790)")) FV["xF"] = txt; else FV["xH"] = txt;
                }
                if (Ps.EqI(rs, "doc")) { matched = true; FV["doc"] = txt.ToUpper(); }
                if (Ps.EqI(rs, "tel")) { matched = true; FV["tel"] = txt; }
                if (Ps.EqI(rs, "nat")) { matched = true; FV["nat"] = txt.ToUpper(); }
                if (rs.StartsWith("x", StringComparison.CurrentCultureIgnoreCase)) { matched = true; FV[rs] = txt; }
                if (Ps.EqI(rs, "price")) { matched = true; double? n = Number(v); if (n.HasValue) price = n; else notes.Add(txt); }
                if (Ps.EqI(rs, "paid")) { matched = true; double? n = Number(v); if (n.HasValue) paid = n; else { payTexts.Add(txt); notes.Add(txt); } }
                if (Ps.EqI(rs, "pend")) { matched = true; double? n = Number(v); if (n.HasValue) pend = n; else notes.Add(txt); }
                if (Ps.EqI(rs, "paytext")) { matched = true; payTexts.Add(txt); notes.Add("Pago: " + txt); }
                if (Ps.EqI(rs, "status")) { matched = true; stTexts.Add(txt); notes.Add(txt); }
                if (Ps.EqI(rs, "obs")) { matched = true; notes.Add(txt); }
                if (rs == "fsol" || rs == "fres")
                {
                    matched = true;
                    object d = Date(v);
                    if (d != null) { FV[rs] = d; if (rs == "fsol" && lay.PresDate) presDate = true; }
                    string rest = PsH.Rep(txt, @"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}", "").Trim(' ', '-');
                    if (d == null || rest.Length > 0) { stTexts.Add(txt); notes.Add(txt); }
                }
                if (rs.StartsWith("#")) { matched = true; notes.Add(rs.Substring(1) + ": " + txt); stTexts.Add(txt); }
                if (!matched)
                {
                    // columna sin cabecera: fecha, estado o nota
                    object d = Date(v);
                    if (d != null && Ps.Blank(FV["fsol"]) && (v is double || PsH.Is(txt, @"^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}$"))) FV["fsol"] = d;
                    else { stTexts.Add(txt); notes.Add(txt); }
                }
            }
            // pago
            foreach (string p in payTexts)
            {
                string k = HandKey(p); double? n = Number(p);
                if (PsH.Is(k, "NO PAGAD|PENDIENTE")) { if (!price.HasValue && n.HasValue) price = n; if (!paid.HasValue) paid = 0.0; }
                else if (PsH.Is(k, "PAGAD")) { if (!price.HasValue && n.HasValue) price = n; if (!paid.HasValue && price.HasValue) paid = price; }
                else if (!price.HasValue && n.HasValue) price = n;
            }
            if (!t.NoStats && !price.HasValue && !paid.HasValue && !pend.HasValue)
            {
                // «255 EUROS PAGADO …» escrito en otra columna
                foreach (string s in notes)
                {
                    Match m = PsH.M(HandKey(s), @"(\d+(?:[.,]\d{1,2})?)\s*(EUROS?|E)?\s*PAGAD");
                    if (m.Success) { price = Ps.ParseMoney(m.Groups[1].Value); paid = price; break; }
                }
            }
            if (!price.HasValue && pend.HasValue) price = (paid.HasValue ? paid.Value : 0.0) + pend.Value;
            FV["price"] = price.HasValue ? (object)price.Value : null;
            if (price.HasValue && paid.HasValue)
            {
                if (paid.Value >= price.Value) { FV["paystate"] = "Pagado"; FV["paid"] = price.Value; }
                else if (paid.Value > 0) { FV["paystate"] = "Parcial"; FV["paid"] = paid.Value; }
                else { FV["paystate"] = "Pendiente"; FV["paid"] = 0.0; }
            }
            else if (price.HasValue) FV["paystate"] = "Pendiente";
            else if (paid.HasValue && paid.Value > 0) { FV["paystate"] = "Parcial"; FV["paid"] = paid.Value; }   // pagó algo, precio desconocido
            FV["status"] = Status(stTexts, presDate);
            if (t.Id == "tasas")
            {
                // orden de las claves de la Hashtable, como en PowerShell
                List<string> parts = new List<string>();
                foreach (object k in row.Keys) if ((int)k != lay.NameCol) parts.Add(Text(row[k]));
                string all = HandKey(string.Join(" ", parts.ToArray()));
                if (PsH.Is(all, "CUENTA PROPIA|POR SU CUENTA")) { FV["tasa"] = "Sí"; FV["xI"] = "Cliente"; }
                else if (PsH.Is(all, "SAFA|POR MI|HECHO")) { FV["tasa"] = "Sí"; FV["xI"] = "Gestoría"; }
                else if (PsH.Is(all, "PAGAD")) FV["tasa"] = "Sí";
            }
            if (notes.Count > 0)
            {
                List<string> uniq = new List<string>();
                foreach (string n in notes) if (!uniq.Contains(n)) uniq.Add(n);     // Select-Object -Unique
                FV["obs"] = string.Join(" · ", uniq.ToArray());
            }
            return FV;
        }

        public static List<ClientRec> FromSheet(Hashtable cells, Tramite t)
        {
            HandLayout lay = Layout(cells, t);
            string prefix = DefaultIdPrefix(t);
            List<ClientRec> list = new List<ClientRec>();
            if (cells.Count == 0) return list;
            int maxRow = 0; foreach (object k in cells.Keys) if ((int)k > maxRow) maxRow = (int)k;
            string tname = HandKey(t.Sheet);
            int gap = 0;
            for (int r = lay.HeaderRow + 1; r <= maxRow; r++)
            {
                Hashtable row = cells[r] as Hashtable;
                string name = ""; if (row != null) name = Text(row[lay.NameCol]);
                if (name.Length == 0) { gap++; if (list.Count > 0 && gap >= 8) break; continue; }
                gap = 0;
                string nk = HandKey(name);
                if (PsH.Is(nk, @"^(PRECIOS?|TOTAL(ES)?|TARIFAS?)\b")) break;            // tabla de precios debajo de los clientes
                if (string.Equals(nk, tname, StringComparison.CurrentCultureIgnoreCase) || (tname.Length > 0 && nk.StartsWith(tname))) continue;   // título
                Dictionary<string, object> FV = FromRow(row, lay, t);
                FV["name"] = name.ToUpper();
                ClientRec c = new ClientRec();
                c.Row = RowBase + r; c.Id = prefix + (r - lay.HeaderRow).ToString().PadLeft(3, '0'); c.F = FV;
                c.Sync = true; c.New = false; c.Updated = ""; c.By = "Excel";
                list.Add(c);
            }
            return list;
        }

        static string Sig(Tramite t, List<ClientRec> list)
        {
            List<ClientRec> l = new List<ClientRec>(list);
            Store.StableSortByRow(l);
            List<string> lines = new List<string>();
            foreach (ClientRec c in l)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(c.Row).Append('|').Append(c.Id).Append('|');
                List<string> vals = new List<string>();
                foreach (Field f in t.Fields) vals.Add(Ps.Str(Store.ToStoreValue(f, Store.Get(c.F, f.Key))));
                sb.Append(string.Join("|", vals.ToArray()));
                lines.Add(sb.ToString());
            }
            return string.Join("\n", lines.ToArray());
        }

        // Sync-HandmadeTabs: lee las pestañas hechas a mano si el Excel ha cambiado; true si cambió algo
        public static bool SyncTabs(bool force)
        {
            string p = Config.Cur.ExcelPath ?? "";
            if (p.Length == 0 || !File.Exists(p)) return false;
            long stamp = File.GetLastWriteTimeUtc(p).Ticks;
            if (!force && HandStamp == stamp) return false;
            List<Tramite> gen = new List<Tramite>(); foreach (Tramite t in Tramite.All) if (t.IsGen) gen.Add(t);
            XlsxBook book;
            try
            {
                book = Xlsx.ReadSheets(p, delegate(string[] all)
                {
                    List<string> w = new List<string>();
                    foreach (Tramite t in gen) w.AddRange(TabCandidates(all, t));
                    return w;
                });
            }
            catch (Exception ex) { App.Log("pestañas propias: no se pudo leer el Excel: " + ex.Message); return false; }
            bool changed = false;
            foreach (Tramite t in gen)
            {
                string src = null;
                foreach (string n in TabCandidates(book.Names, t))
                    if (book.Sheets.ContainsKey(n) && !CellsGestem((Hashtable)book.Sheets[n], t)) { src = n; break; }
                List<ClientRec> recs = new List<ClientRec>();
                if (src != null) { recs = FromSheet((Hashtable)book.Sheets[src], t); HandTabs[t.Id] = src; } else HandTabs.Remove(t.Id);
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    if (s == null && recs.Count == 0) continue;
                    if (s == null) s = NewEmptyStore(t, null);
                    // clientes que ya se editaron en GESTEM: ahora son de GESTEM, no se vuelven a leer
                    List<string> adoptedKeys = new List<string>(); foreach (string a in s.Adopted) adoptedKeys.Add(HandKey(a));
                    if (adoptedKeys.Count > 0)
                    {
                        List<ClientRec> f2 = new List<ClientRec>();
                        foreach (ClientRec c in recs) if (!PsH.ContainsI(adoptedKeys, HandKey(Ps.Str(Store.Get(c.F, "name"))))) f2.Add(c);
                        recs = f2;
                    }
                    List<ClientRec> old = new List<ClientRec>(); foreach (ClientRec c in s.Clients) if (c.Row >= RowBase) old.Add(c);
                    if (Sig(t, old) != Sig(t, recs))
                    {
                        List<ClientRec> keep = new List<ClientRec>();
                        foreach (ClientRec c in s.Clients) if (c.Row < RowBase) keep.Add(c);
                        keep.AddRange(recs);
                        s.Clients = keep;
                        Store.Write(t, s);
                        changed = true;
                        if (src != null) App.Log(string.Format("leído del Excel (pestaña «{0}»): {1} clientes", src, recs.Count));
                    }
                }
                finally { Store.ExitLock(lk); }
            }
            HandStamp = stamp;
            return changed;
        }

        // Update-ListsFromFile: 'ok' (listas al día sin Excel), 'skip' o 'excel' (hace falta Excel)
        public static string UpdateListsFromFile(Tramite t, Hashtable dv)
        {
            List<string> need = new List<string>();
            foreach (Field f in t.Fields) if ((f.List ?? "").StartsWith("excel:", StringComparison.OrdinalIgnoreCase)) { string n = f.List.Substring(6); if (!need.Contains(n)) need.Add(n); }
            if (need.Count == 0) return "skip";
            List<string> names = new List<string>(); foreach (object v in dv.Values) names.Add(((DvSheet)v).Name);
            string[] cands = new string[] { t.Sheet };
            if (t.IsGen) cands = TabCandidates(names.ToArray(), t);
            bool range = false;
            foreach (string c in cands)
            {
                string k = Xlsx.SheetKey(c);
                if (!dv.ContainsKey(k)) continue;
                Hashtable cols = ((DvSheet)dv[k]).Cols;
                List<string> hits = new List<string>(); foreach (string n in need) if (cols.ContainsKey(n)) hits.Add(n);
                bool anyRange = false; foreach (string h in hits) if (cols[h] is string) anyRange = true;
                if (anyRange) { range = true; continue; }
                if (hits.Count < need.Count) continue;
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    if (s == null) return "skip";
                    foreach (string n in need) s.Meta.Lists["excel:" + n] = (string[])cols[n];
                    Store.Write(t, s);
                }
                finally { Store.ExitLock(lk); }
                return "ok";
            }
            return range ? "excel" : "skip";
        }
    }
}

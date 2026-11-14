// Escribe los clientes directamente en el .xlsx (sin abrir Excel). Si algo no es seguro devuelve null
// y GESTEM usa Excel como siempre. Copiado de Get-ClientOps / Save-ClientsToFile / Backup-Excel.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Gestem
{
    public class ClientOp { public string Col, Kind = "value", Type = "Text"; public object Value; }
    public class FileItem { public List<ClientOp> Ops; public int? Row; public string Expect; public bool IsNew; }
    public class FileResult { public string Id; public int Row; }

    public static class ExcelFile
    {
        const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static int SumFixed;

        // Get-ClientOps
        public static List<ClientOp> ClientOps(Tramite T, Dictionary<string, object> vals, Dictionary<string, double> tariff)
        {
            string tipo = Ps.Str(Store.Get(vals, "tipo"));
            List<ClientOp> ops = new List<ClientOp>();
            foreach (Field f in T.Fields)
            {
                object v = Store.Get(vals, f.Key);
                ClientOp op = new ClientOp(); op.Col = f.Col;
                string role = f.Role ?? "";
                if (Ps.EqI(role, "nationality"))
                {
                    if (Ps.Str(v).Length == 0 || Ps.EqI(Ps.Str(v), T.DefaultNationality)) { op.Kind = "formula"; op.Value = T.DefaultNationality; }
                    else op.Value = Ps.Str(v);
                }
                else if (Ps.EqI(role, "price"))
                {
                    double? tar = null; double? pv = Ps.Num(v);
                    if (tipo.Length > 0 && tariff != null && tariff.ContainsKey(tipo)) tar = tariff[tipo];
                    if (!pv.HasValue || (tar.HasValue && pv.Value == tar.Value)) { op.Kind = "formula"; op.Value = pv.HasValue ? (object)pv.Value : (tar.HasValue ? (object)tar.Value : null); op.Type = "Money"; }
                    else { op.Value = pv.Value; op.Type = "Money"; }
                }
                else if (Ps.EqI(role, "paid"))
                {
                    double? pv = Ps.Num(v);
                    if (Ps.Str(Store.Get(vals, "paystate")) == "Parcial" && pv.HasValue) { op.Value = pv.Value; op.Type = "Money"; }
                    else { op.Kind = "formula"; op.Value = pv.HasValue ? (object)pv.Value : null; op.Type = "Money"; }
                }
                else
                {
                    if (f.Type == "Date") { double? d = Ps.CellDate(v); op.Value = d.HasValue ? (object)d.Value : null; op.Type = "Date"; }
                    else if (f.Type == "Money") { double? n = Ps.Num(v); op.Value = n.HasValue ? (object)n.Value : null; op.Type = "Money"; }
                    else if (v != null && Ps.Str(v) != "") op.Value = Ps.Str(v);
                }
                ops.Add(op);
            }
            return ops;
        }

        // Backup-Excel: copia con fecha y hora; se quedan las 40 últimas
        public static void BackupExcel()
        {
            try
            {
                string src = Config.Cur.ExcelPath;
                string name = Path.GetFileNameWithoutExtension(src);
                string dst = Path.Combine(App.BackupDir, name + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".xlsx");
                File.Copy(src, dst, true);
                List<FileInfo> all = new List<FileInfo>(new DirectoryInfo(App.BackupDir).GetFiles("*.xlsx"));
                all.Sort(delegate(FileInfo a, FileInfo b) { return b.LastWriteTime.CompareTo(a.LastWriteTime); });
                for (int i = 40; i < all.Count; i++) { try { all[i].Delete(); } catch { } }
            }
            catch (Exception ex) { App.Log("copia de seguridad fallida: " + ex.Message); }
        }

        static XmlDocument ReadXml(ZipArchive zip, string name)
        {
            ZipArchiveEntry e = zip.GetEntry(name);
            if (e == null) return null;
            using (StreamReader sr = new StreamReader(e.Open()))
            {
                XmlDocument x = new XmlDocument(); x.PreserveWhitespace = true; x.LoadXml(sr.ReadToEnd()); return x;
            }
        }
        static void WriteXml(ZipArchive zip, string name, XmlDocument x)
        {
            ZipArchiveEntry old = zip.GetEntry(name); if (old != null) old.Delete();
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (StreamWriter sw = new StreamWriter(e.Open(), new UTF8Encoding(false))) x.Save(sw);
        }

        // calcChain.xml es la lista de Excel de las celdas con fórmula. Si una de esas celdas ya no tiene fórmula,
        // Excel da el archivo por dañado y no lo abre. Se quita la lista entera: Excel la vuelve a crear sola.
        public static bool RemoveCalcChain(ZipArchive zip)
        {
            ZipArchiveEntry cc = zip.GetEntry("xl/calcChain.xml");
            if (cc == null) return false;
            cc.Delete();
            XmlDocument rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
            if (rels != null)
            {
                List<XmlNode> kill = new List<XmlNode>();
                foreach (XmlNode r in rels.DocumentElement.ChildNodes) if (r.LocalName == "Relationship" && ((XmlElement)r).GetAttribute("Type").EndsWith("/calcChain")) kill.Add(r);
                foreach (XmlNode r in kill) rels.DocumentElement.RemoveChild(r);
                WriteXml(zip, "xl/_rels/workbook.xml.rels", rels);
            }
            XmlDocument ct = ReadXml(zip, "[Content_Types].xml");
            if (ct != null)
            {
                List<XmlNode> kill = new List<XmlNode>();
                foreach (XmlNode o in ct.DocumentElement.ChildNodes) if (o.LocalName == "Override" && ((XmlElement)o).GetAttribute("PartName") == "/xl/calcChain.xml") kill.Add(o);
                foreach (XmlNode o in kill) ct.DocumentElement.RemoveChild(o);
                WriteXml(zip, "[Content_Types].xml", ct);
            }
            return true;
        }

        static readonly System.Text.RegularExpressions.Regex BadXmlNs = new System.Text.RegularExpressions.Regex(@"xmlns:(\w+)=""http://www\.w3\.org/XML/1998/namespace""", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // 1.7.2 y 1.7.3 podían dejar el .xlsx de forma que Excel no lo abre (calcChain.xml y «d1p1:space»):
        // una vez por PC, con el archivo cerrado, se arreglan las dos cosas
        public static void RepairExcelFile()
        {
            if (Config.Cur.ExcelRepaired) return;
            string path = Config.Cur.ExcelPath ?? "";
            if (path.Length == 0 || !File.Exists(path)) return;
            try { using (FileStream probe = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } }
            catch { App.Log("reparar Excel: el archivo está abierto, se hará otro día"); return; }
            string tmp = Path.Combine(Path.GetTempPath(), "GESTEM_reparar_" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(path, tmp, true);
            try
            {
                bool removed;
                using (ZipArchive zip = ZipFile.Open(tmp, ZipArchiveMode.Update))
                {
                    removed = RemoveCalcChain(zip);
                    List<ZipArchiveEntry> sheets = new List<ZipArchiveEntry>();
                    foreach (ZipArchiveEntry en in zip.Entries) if (en.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && en.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && en.FullName.IndexOf('/', 14) < 0) sheets.Add(en);
                    foreach (ZipArchiveEntry en in sheets)
                    {
                        string name = en.FullName; string txt;
                        using (StreamReader sr = new StreamReader(en.Open(), Encoding.UTF8)) txt = sr.ReadToEnd();
                        List<string> bad = new List<string>();
                        foreach (System.Text.RegularExpressions.Match m in BadXmlNs.Matches(txt)) if (!bad.Contains(m.Groups[1].Value)) bad.Add(m.Groups[1].Value);
                        if (bad.Count == 0) continue;
                        foreach (string px in bad)
                        {
                            txt = txt.Replace(" xmlns:" + px + "=\"http://www.w3.org/XML/1998/namespace\"", "");
                            txt = txt.Replace(px + ":space=", "xml:space=");
                        }
                        en.Delete();
                        ZipArchiveEntry ne = zip.CreateEntry(name, CompressionLevel.Optimal);
                        using (StreamWriter sw = new StreamWriter(ne.Open(), new UTF8Encoding(false))) sw.Write(txt);
                        removed = true;
                        App.Log("reparar Excel: " + name + " (" + bad.Count + " espacio(s) de nombres no válidos)");
                    }
                }
                if (removed)
                {
                    using (ZipArchive chk = ZipFile.OpenRead(tmp))
                    {
                        if (chk.GetEntry("xl/workbook.xml") == null || chk.GetEntry("xl/calcChain.xml") != null) throw new GestemException("comprobación fallida");
                        foreach (ZipArchiveEntry en in chk.Entries)
                        {
                            if (!(en.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) && en.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))) continue;
                            string t2; using (StreamReader sr = new StreamReader(en.Open())) { t2 = sr.ReadToEnd(); XmlDocument xx = new XmlDocument(); xx.LoadXml(t2); }
                            if (BadXmlNs.IsMatch(t2)) throw new GestemException("comprobación fallida");
                        }
                    }
                    BackupExcel();
                    File.Copy(tmp, path, true);
                    App.Log("reparar Excel: archivo arreglado (Excel vuelve a crear calcChain.xml)");
                }
                Config.Cur.ExcelRepaired = true; Config.Save();
            }
            catch (Exception e) { App.Log("reparar Excel: " + e.Message); }
            finally { try { File.Delete(tmp); } catch { } }
        }

        class Ctx
        {
            public XmlDocument Doc; public Hashtable Rows; public XmlElement Sd; public List<string> Shared; public Tramite T;
            public Hashtable ColFx = PsH.New();

            public string CellText(XmlElement c)
            {
                if (c == null) return "";
                string tt = c.GetAttribute("t");
                foreach (XmlNode ch in c.ChildNodes)
                {
                    if (ch.LocalName == "is") return ch.InnerText;
                    if (ch.LocalName == "v")
                    {
                        if (tt == "s") { int i = int.Parse(ch.InnerText, CultureInfo.InvariantCulture); if (i < Shared.Count) return Shared[i]; return ""; }
                        return ch.InnerText;
                    }
                }
                return "";
            }

            public XmlElement GetCell(int r, string col, bool create)
            {
                XmlElement rw = Rows[r] as XmlElement;
                if (rw == null)
                {
                    if (!create) return null;
                    rw = Doc.CreateElement("row", NS); rw.SetAttribute("r", r.ToString(CultureInfo.InvariantCulture));
                    XmlElement after = null;
                    foreach (object k in Rows.Keys) if ((int)k < r && (after == null || (int)k > int.Parse(after.GetAttribute("r")))) after = (XmlElement)Rows[k];
                    if (after != null) Sd.InsertAfter(rw, after); else Sd.PrependChild(rw);
                    Rows[r] = rw;
                }
                string rf = col + r; int n = Xlsx.ColNum(col); XmlElement before = null;
                foreach (XmlNode cn in rw.ChildNodes)
                {
                    if (cn.LocalName != "c") continue;
                    XmlElement c = (XmlElement)cn;
                    string cr = c.GetAttribute("r");
                    if (cr == rf) return c;
                    if (before == null) { System.Text.RegularExpressions.Match m = PsH.M(cr, "^([A-Z]+)"); if (m.Success && Xlsx.ColNum(m.Groups[1].Value) > n) before = c; }
                }
                if (!create) return null;
                XmlElement nc = Doc.CreateElement("c", NS); nc.SetAttribute("r", rf);
                if (before != null) rw.InsertBefore(nc, before); else rw.AppendChild(nc);
                return nc;
            }

            // formato de esa columna en otra fila (para fechas en celdas nuevas)
            public string ColStyle(string col)
            {
                List<int> keys = new List<int>(); foreach (object k in Rows.Keys) keys.Add((int)k); keys.Sort();
                foreach (int k in keys)
                {
                    if (k < T.FirstRow) continue;
                    foreach (XmlNode c in ((XmlElement)Rows[k]).ChildNodes)
                        if (c.LocalName == "c" && ((XmlElement)c).GetAttribute("r") == col + k && ((XmlElement)c).GetAttribute("s").Length > 0) return ((XmlElement)c).GetAttribute("s");
                }
                return null;
            }

            public static bool IsSum(XmlNode fx) { return fx != null && (fx.InnerText ?? "").Trim().TrimStart('=').ToUpperInvariant().StartsWith("SUM("); }

            // ¿esa columna tiene fórmulas en OTRAS filas de clientes? (filas con nombre; nunca la fila de totales)
            public bool ColHasFormula(string col, int skip)
            {
                if (ColFx.ContainsKey(col)) return (bool)ColFx[col];
                ColFx[col] = false;
                foreach (object ko in Rows.Keys)
                {
                    int k = (int)ko;
                    if (k < T.FirstRow || k > T.LastRow || k == skip) continue;
                    if (CellText(GetCell(k, T.NameCol, false)).Trim().Length == 0) continue;
                    foreach (XmlNode c in ((XmlElement)Rows[k]).ChildNodes)
                        if (c.LocalName == "c" && ((XmlElement)c).GetAttribute("r") == col + k)
                            foreach (XmlNode ch in c.ChildNodes) if (ch.LocalName == "f" && !IsSum(ch)) { ColFx[col] = true; return true; }
                }
                return false;
            }
        }

        public static FileResult SaveOne(Tramite T, List<ClientOp> ops, int? row, string expect, bool isNew)
        {
            FileItem it = new FileItem(); it.Ops = ops; it.Row = row; it.Expect = expect; it.IsNew = isNew;
            List<FileResult> r = SaveMany(T, new List<FileItem>(new FileItem[] { it }), false);
            return r != null && r.Count > 0 ? r[0] : null;
        }

        // Save-ClientsToFile: varios clientes del mismo trámite en una sola pasada por el archivo
        // onlyBroken: solo arregla celdas con una fórmula SUM(...) copiada por error de la fila de totales
        public static List<FileResult> SaveMany(Tramite T, List<FileItem> items, bool onlyBroken)
        {
            SumFixed = 0;
            string path = Config.Cur.ExcelPath ?? "";
            if (path.Length == 0 || !File.Exists(path)) return null;
            // abierto en Excel (aquí o en otro PC): Excel tiene el archivo bloqueado -> se usa Excel
            try { using (FileStream probe = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } }
            catch { App.Log("Excel directo: el archivo está abierto en otro programa"); return null; }
            string tmp = Path.Combine(Path.GetTempPath(), "GESTEM_escritura_" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(path, tmp, true);
            List<FileResult> result = null;
            try
            {
                string sheetPath = null;
                using (ZipArchive zip = ZipFile.Open(tmp, ZipArchiveMode.Update))
                {
                    List<string> shared = new List<string>();
                    XmlDocument sst = ReadXml(zip, "xl/sharedStrings.xml");
                    if (sst != null) foreach (XmlNode si in sst.DocumentElement.ChildNodes) if (si.LocalName == "si") shared.Add(si.InnerText);
                    XmlDocument wbx = ReadXml(zip, "xl/workbook.xml"); XmlDocument rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
                    Hashtable targets = PsH.New();
                    foreach (XmlNode r in rels.DocumentElement.ChildNodes) if (r.LocalName == "Relationship") targets[((XmlElement)r).GetAttribute("Id")] = ((XmlElement)r).GetAttribute("Target");
                    List<string[]> sheets = new List<string[]>();
                    foreach (XmlElement s in wbx.GetElementsByTagName("sheet", NS))
                    {
                        string p = Ps.Str(targets[s.GetAttribute("id", NSR)]).TrimStart('/'); if (!p.StartsWith("xl/")) p = "xl/" + p;
                        sheets.Add(new string[] { s.GetAttribute("name"), p });
                    }
                    List<string> names = new List<string>(); foreach (string[] s in sheets) names.Add(s[0]);
                    List<string> cands = new List<string>();
                    if (T.IsGen) cands.AddRange(Hand.TabCandidates(names.ToArray(), T));
                    else foreach (string n in names) if (Ps.EqI(Xlsx.SheetKey(n), Xlsx.SheetKey(T.Sheet))) cands.Add(n);
                    Ctx cx = null;
                    foreach (string cand in cands)
                    {
                        string[] sh = null; foreach (string[] s in sheets) if (Ps.EqI(s[0], cand)) { sh = s; break; }
                        if (sh == null) continue;
                        XmlDocument x = ReadXml(zip, sh[1]);
                        if (x == null) continue;
                        XmlNodeList sdl = x.GetElementsByTagName("sheetData", NS);
                        if (sdl.Count == 0) continue;
                        XmlElement sdn = (XmlElement)sdl[0];
                        Hashtable rmap = PsH.New();
                        foreach (XmlNode rw in sdn.ChildNodes) if (rw.LocalName == "row") rmap[int.Parse(((XmlElement)rw).GetAttribute("r"), CultureInfo.InvariantCulture)] = rw;
                        Ctx tc = new Ctx(); tc.Doc = x; tc.Rows = rmap; tc.Sd = sdn; tc.Shared = shared; tc.T = T;
                        if (T.IsGen)
                        {
                            // solo una pestaña con el formato de GESTEM (cabecera «Cliente»/«Nombre…» encima de la primera fila)
                            XmlElement hr = rmap[T.FirstRow - 1] as XmlElement; XmlElement hc = null;
                            if (hr != null) foreach (XmlNode c in hr.ChildNodes) if (c.LocalName == "c" && ((XmlElement)c).GetAttribute("r") == T.NameCol + (T.FirstRow - 1)) hc = (XmlElement)c;
                            string hk = Xlsx.SheetKey(tc.CellText(hc));
                            if (!(hk.StartsWith("CLIENTE") || hk.StartsWith("NOMBRE"))) continue;
                        }
                        cx = tc; sheetPath = sh[1]; break;
                    }
                    if (cx == null) { App.Log("Excel directo: no hay pestaña de GESTEM para " + T.Sheet); return null; }   // crear la pestaña lo hace Excel
                    List<FileResult> results = new List<FileResult>();
                    bool changed = false;
                    foreach (FileItem it in items)
                    {
                        // fila: la guardada, salvo que en el Excel ya haya otro cliente escrito a mano
                        int? row = it.Row;
                        if (row.HasValue && row.Value != 0 && it.IsNew)
                        {
                            string cur = cx.CellText(cx.GetCell(row.Value, T.NameCol, false)).Trim();
                            if (cur.Length > 0 && cur != (it.Expect ?? "").Trim()) row = null;
                        }
                        if (!row.HasValue || row.Value == 0)
                        {
                            row = null;
                            for (int r = T.FirstRow; r <= T.LastRow; r++) if (cx.CellText(cx.GetCell(r, T.NameCol, false)).Trim().Length == 0) { row = r; break; }
                            if (!row.HasValue) throw new GestemException(Tx.L("errFull"));
                        }
                        int rr = row.Value;
                        foreach (ClientOp op in it.Ops)
                        {
                            XmlElement c = cx.GetCell(rr, op.Col, false);
                            bool empty = op.Value == null || Ps.Str(op.Value) == "";
                            if (c == null && empty) continue;
                            XmlNode fx = null; if (c != null) foreach (XmlNode ch in c.ChildNodes) if (ch.LocalName == "f") fx = ch;
                            bool broken = Ctx.IsSum(fx);     // SUM(...) en una fila de cliente: copiada por error de la fila de totales
                            if (onlyBroken && !broken) continue;
                            if (broken) { fx = null; SumFixed++; }
                            if (op.Kind == "formula" && fx != null) continue;                              // la fórmula del Excel se queda
                            if (op.Kind == "formula" && cx.ColHasFormula(op.Col, rr)) { App.Log("Excel directo: la columna " + op.Col + " tiene fórmula"); return null; }
                            if (fx != null && ((XmlElement)fx).GetAttribute("t") == "shared" && ((XmlElement)fx).GetAttribute("ref").Length > 0) { App.Log("Excel directo: fórmula compartida en " + op.Col + rr); return null; }
                            if (c == null) c = cx.GetCell(rr, op.Col, true);
                            if (op.Type == "Date" && !empty && c.GetAttribute("s").Length == 0)
                            {
                                string st = cx.ColStyle(op.Col);
                                if (st == null) { App.Log("Excel directo: la columna " + op.Col + " no tiene formato de fecha"); return null; }
                                c.SetAttribute("s", st);
                            }
                            List<XmlNode> kill = new List<XmlNode>();
                            foreach (XmlNode ch in c.ChildNodes) if (ch.LocalName == "f" || ch.LocalName == "v" || ch.LocalName == "is") kill.Add(ch);
                            foreach (XmlNode ch in kill) c.RemoveChild(ch);
                            changed = true;
                            c.RemoveAttribute("t");
                            if (empty) continue;
                            if (op.Type == "Date" || op.Type == "Money")
                            {
                                XmlElement v = cx.Doc.CreateElement("v", NS); v.InnerText = Convert.ToDouble(op.Value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture);
                                c.AppendChild(v);
                            }
                            else
                            {
                                c.SetAttribute("t", "inlineStr");
                                XmlElement isn = cx.Doc.CreateElement("is", NS); XmlElement tx = cx.Doc.CreateElement("t", NS);
                                string sv = Ps.Str(op.Value);
                                tx.InnerText = sv;
                                if (sv != sv.Trim())
                                {
                                    // xml:space="preserve" (con SetAttribute(..., espacio de nombres) salía «d1p1:space» y Excel daba el archivo por dañado)
                                    XmlAttribute sa = cx.Doc.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace"); sa.Value = "preserve"; tx.Attributes.Append(sa);
                                }
                                isn.AppendChild(tx); c.AppendChild(isn);
                            }
                        }
                        FileResult fr = new FileResult(); fr.Id = cx.CellText(cx.GetCell(rr, T.IdCol, false)).Trim(); fr.Row = rr;
                        results.Add(fr);
                    }
                    if (!changed) { result = results; sheetPath = null; }
                    else
                    {
                        WriteXml(zip, sheetPath, cx.Doc);
                        RemoveCalcChain(zip);
                        // que Excel vuelva a calcular los totales al abrir
                        XmlNodeList calc = wbx.GetElementsByTagName("calcPr", NS);
                        if (calc.Count > 0) { ((XmlElement)calc[0]).SetAttribute("fullCalcOnLoad", "1"); WriteXml(zip, "xl/workbook.xml", wbx); }
                        result = results;
                    }
                }
                if (result == null) return null;
                if (sheetPath == null) return result;      // nada que cambiar en el archivo
                // comprobación: el archivo nuevo se abre y la pestaña se lee bien
                using (ZipArchive chk = ZipFile.OpenRead(tmp))
                {
                    string txt;
                    using (StreamReader sr = new StreamReader(chk.GetEntry(sheetPath).Open())) { txt = sr.ReadToEnd(); XmlDocument xx = new XmlDocument(); xx.LoadXml(txt); }
                    if (BadXmlNs.IsMatch(txt)) throw new GestemException("XML no válido para Excel");
                    if (chk.GetEntry("xl/calcChain.xml") != null) throw new GestemException("calcChain.xml sigue en el archivo");
                }
                if (Config.Cur.Backup) BackupExcel();
                File.Copy(tmp, path, true);      // si Excel lo acaba de abrir, falla aquí y se usa Excel
                if (onlyBroken) App.Log(string.Format("Excel (directo): {0} celda(s) con SUM(...) copiada de la fila de totales, arreglada(s) · {1}", SumFixed, T.Sheet));
                else
                {
                    List<string> parts = new List<string>(); foreach (FileResult r in result) parts.Add("fila " + r.Row + " (" + r.Id + ")");
                    App.Log("Excel (directo): " + string.Join(", ", parts.ToArray()));
                }
                return result;
            }
            finally { try { File.Delete(tmp); } catch { } }
        }
    }
}

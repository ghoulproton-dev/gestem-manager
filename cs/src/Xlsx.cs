// Lee el .xlsx directamente (aunque esté abierto en Excel): pestañas y listas desplegables.
// Copiado de Read-XlsxSheets / Read-XlsxValidationLists.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Xml;

namespace Gestem
{
    public class XlsxBook
    {
        public string[] Names = new string[0];
        public Hashtable Sheets = PsH.New();        // nombre -> celdas (Hashtable fila -> Hashtable columna -> valor)
    }

    public class DvSheet { public string Name; public Hashtable Cols = PsH.New(); }   // columna -> string[] o "#RANGE"

    public static class Xlsx
    {
        public static int ColNum(string c) { int n = 0; foreach (char ch in c.ToUpperInvariant()) n = n * 26 + (ch - 64); return n; }

        // nombre de pestaña sin tildes, mayúsculas ni espacios de más
        public static string SheetKey(string s)
        {
            string d = (s ?? "").Trim().ToUpperInvariant().Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();
            foreach (char ch in d) if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark) sb.Append(ch);
            return PsH.Rep(sb.ToString(), @"\s+", " ");
        }

        static XmlDocument Load(ZipArchive zip, string n)
        {
            ZipArchiveEntry e = zip.GetEntry(n);
            if (e == null) return null;
            using (StreamReader sr = new StreamReader(e.Open()))
            {
                XmlDocument x = new XmlDocument();
                x.LoadXml(sr.ReadToEnd());
                return x;
            }
        }

        class SheetRef { public string Name, Path; }

        static List<SheetRef> SheetList(ZipArchive zip)
        {
            XmlDocument wbx = Load(zip, "xl/workbook.xml");
            XmlDocument rels = Load(zip, "xl/_rels/workbook.xml.rels");
            Hashtable target = PsH.New();
            foreach (XmlElement r in rels.SelectNodes("//*[local-name()='Relationship']")) target[r.GetAttribute("Id")] = r.GetAttribute("Target");
            List<SheetRef> sheets = new List<SheetRef>();
            foreach (XmlElement sh in wbx.SelectNodes("//*[local-name()='sheet']"))
            {
                string rid = null;
                foreach (XmlAttribute a in sh.Attributes) if (a.LocalName == "id") { rid = a.Value; break; }
                SheetRef s = new SheetRef();
                s.Name = sh.GetAttribute("name");
                s.Path = rid != null && target[rid] != null ? (string)target[rid] : "";
                sheets.Add(s);
            }
            return sheets;
        }
        static string EntryPath(string p) { p = p.TrimStart('/'); if (!p.StartsWith("xl/")) p = "xl/" + p; return p; }

        public static XlsxBook ReadSheets(string path, Func<string[], IEnumerable<string>> pick)
        {
            XlsxBook book = new XlsxBook();
            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                List<string> shared = new List<string>();
                XmlDocument ss = Load(zip, "xl/sharedStrings.xml");
                if (ss != null)
                    foreach (XmlNode si in ss.SelectNodes("//*[local-name()='si']"))
                    {
                        StringBuilder b = new StringBuilder();
                        foreach (XmlNode t in si.SelectNodes(".//*[local-name()='t']")) b.Append(t.InnerText);
                        shared.Add(b.ToString());
                    }
                List<SheetRef> sheets = SheetList(zip);
                List<string> all = new List<string>();
                foreach (SheetRef s in sheets) all.Add(s.Name);
                book.Names = all.ToArray();
                List<string> want = new List<string>(pick(book.Names));
                foreach (SheetRef s in sheets)
                {
                    if (!PsH.ContainsI(want, s.Name)) continue;
                    ZipArchiveEntry e = zip.GetEntry(EntryPath(s.Path));
                    if (e == null) continue;
                    book.Sheets[s.Name] = ReadCells(e, shared);
                }
            }
            return book;
        }

        // lectura en flujo: las pestañas tienen miles de celdas vacías con formato; solo se miran las que tienen valor
        static Hashtable ReadCells(ZipArchiveEntry e, List<string> shared)
        {
            Hashtable cells = PsH.New();
            using (XmlReader xr = XmlReader.Create(e.Open()))
            {
                string rf = null, ct = null, buf = null; bool inIs = false;
                bool more = xr.Read();
                while (more)
                {
                    bool next = true;
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        switch (xr.LocalName)
                        {
                            case "c": rf = xr.GetAttribute("r"); ct = xr.GetAttribute("t"); buf = null; if (xr.IsEmptyElement) rf = null; break;
                            case "v": if (rf != null) { buf = xr.ReadElementContentAsString(); next = false; } break;
                            case "is": inIs = true; buf = ""; break;
                            case "t": if (inIs && rf != null) { buf += xr.ReadElementContentAsString(); next = false; } break;
                        }
                    }
                    else if (xr.NodeType == XmlNodeType.EndElement)
                    {
                        if (xr.LocalName == "is") inIs = false;
                        else if (xr.LocalName == "c" && rf != null)
                        {
                            System.Text.RegularExpressions.Match m = PsH.M(rf, @"^([A-Z]+)(\d+)$");
                            if (buf != null && m.Success)
                            {
                                int col = ColNum(m.Groups[1].Value); int row = int.Parse(m.Groups[2].Value);
                                if (row > 10 && cells.ContainsKey(10) && ((Hashtable)cells[10]).ContainsKey(2))
                                {
                                    // pestaña con el formato de GESTEM (cabecera en B10): no hace falta leer más
                                    string h10 = SheetKey(Ps.Str(((Hashtable)cells[10])[2]));
                                    if (h10.StartsWith("CLIENTE") || h10.StartsWith("NOMBRE")) break;
                                }
                                object v = null;
                                if (ct == "s") v = shared[int.Parse(buf, CultureInfo.InvariantCulture)];
                                else if (ct == "str" || ct == "inlineStr") v = buf;
                                else if (ct == "b" || ct == "e") v = null;
                                else { double d; if (double.TryParse(buf, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) v = d; else v = buf; }
                                if (v != null && Ps.Str(v).Trim() != "")
                                {
                                    if (!cells.ContainsKey(row)) cells[row] = PsH.New();
                                    ((Hashtable)cells[row])[col] = v;
                                }
                            }
                            rf = null;
                        }
                    }
                    if (next) more = xr.Read(); else more = !xr.EOF;
                }
            }
            return cells;
        }

        // listas desplegables de todas las pestañas (clave = SheetKey del nombre)
        public static Hashtable ReadValidationLists(string path)
        {
            Hashtable outp = PsH.New();
            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (SheetRef sh in SheetList(zip))
                {
                    ZipArchiveEntry e = zip.GetEntry(EntryPath(sh.Path));
                    if (e == null) continue;
                    DvSheet dv = new DvSheet(); dv.Name = sh.Name;
                    using (XmlReader xr = XmlReader.Create(e.Open()))
                    {
                        string sq = null; bool isList = false;
                        while (xr.Read())
                        {
                            if (xr.NodeType != XmlNodeType.Element) continue;
                            if (xr.LocalName == "dataValidation") { sq = xr.GetAttribute("sqref"); isList = xr.GetAttribute("type") == "list"; }
                            else if (xr.LocalName == "formula1" && isList && sq != null)
                            {
                                string f = (xr.ReadInnerXml() ?? "").Trim();
                                object val = "#RANGE";
                                if (f.StartsWith("&quot;")) f = WebUtility.HtmlDecode(f);
                                if (f.StartsWith("\"") && !f.Contains("<"))
                                {
                                    // un trozo que empieza por espacio es parte de la opción anterior
                                    List<string> items = new List<string>();
                                    foreach (string piece in f.Trim('"').Split(new char[] { ',', ';' }))
                                    {
                                        string v = piece.Trim().Trim('"').Trim();
                                        if (v.Length == 0) continue;
                                        if (piece.StartsWith(" ") && items.Count > 0) items[items.Count - 1] = items[items.Count - 1] + ", " + v; else items.Add(v);
                                    }
                                    val = items.ToArray();
                                }
                                foreach (string rg in sq.Split(' '))
                                {
                                    System.Text.RegularExpressions.Match m = PsH.M(rg, @"^\$?([A-Z]+)");
                                    if (m.Success && !dv.Cols.ContainsKey(m.Groups[1].Value)) dv.Cols[m.Groups[1].Value] = val;
                                }
                                isList = false;
                            }
                        }
                    }
                    outp[SheetKey(sh.Name)] = dv;
                }
            }
            return outp;
        }
    }
}

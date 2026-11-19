// Etapa 3d: carpeta de documentos de cada cliente («GESTEM documentos\<pestaña>\<nº> - <nombre>»).
// Copiado de Get-SafeName, Get-DocsRoot, Get-LocalDocsRoot, Get-TramiteDocsDir, Get-ClientDocsDir, Merge-DocsFolder,
// Test-SameNameKey, Repair-DocsFolders, Repair-AllDocsFolders, Update-DocsMap y Get-DocsCount de GESTEM.ps1.
// shortcut: elegir la carpeta en Ajustes (Set-DocsRoot / Choose-DocsFolder) es de la etapa 4; buscarla sola en la nube
// (Initialize-DocsFolder / Find-CloudDocsDir) de la 5b, y la copia a la carpeta externa (Sync-DocsBackup) de la 5.
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace Gestem
{
    public class MergeResult { public int Moved, Same, Renamed, Failed; }

    public static class Docs
    {
        public static Hashtable Map = PsH.New();    // $script:DocsMap: nº de expediente -> nº de documentos (trámite actual)

        public static string SafeName(string s) { return PsH.Rep(PsH.Rep(Ps.Str(s), "[\\\\/:*?\"<>|]", " "), @"\s+", " ").Trim().TrimEnd('.'); }

        // la carpeta elegida en Ajustes (o encontrada sola en la nube); si no, la de al lado del Excel
        public static string Root() { string d = Ps.Str(Config.Cur.DocsDir); return d.Length > 0 ? d : LocalRoot(); }
        public static string LocalRoot()
        {
            string x = Ps.Str(Config.Cur.ExcelPath);
            if (x.Length == 0) return null;
            return Path.Combine(Path.GetDirectoryName(x), "GESTEM documentos");
        }
        public static string TramiteDir(Tramite t) { string r = Root(); return r == null ? null : Path.Combine(r, SafeName(t.Sheet)); }

        static string IdOf(string id, int row) { string s = SafeName(id); return s.Length > 0 ? s : "fila " + row; }

        public static string ClientDir(Tramite t, ClientRow r, bool create)
        {
            string tdir = TramiteDir(t);
            if (tdir == null) return null;
            string id = IdOf(r.Id, r.Row);
            string p = Path.Combine(tdir, id + " - " + SafeName(r.Name));
            if (Directory.Exists(p)) return p;    // mismo nº y mismo nombre (puede haber otra carpeta con el mismo nº de otro PC)
            if (Directory.Exists(tdir))
                foreach (string d in Directory.GetDirectories(tdir))
                {
                    string n = Path.GetFileName(d);
                    if (n == id || n.StartsWith(id + " - ", StringComparison.Ordinal)) return d;
                }
            if (create) Directory.CreateDirectory(p);
            return p;
        }

        static string Sha(string path)
        {
            using (SHA256 h = SHA256.Create()) using (FileStream f = File.OpenRead(path)) return Convert.ToBase64String(h.ComputeHash(f));
        }

        // pasa los documentos de from a to sin sobrescribir nunca: mismo nombre y contenido -> ya estaba (se quita el de from);
        // mismo nombre y distinto contenido -> se guardan los dos («nombre (2).pdf»). copy: copia y deja from como estaba
        public static MergeResult Merge(string from, string to, bool copy)
        {
            MergeResult res = new MergeResult();
            if (string.IsNullOrEmpty(from) || !Directory.Exists(from)) return res;
            string src = Path.GetFullPath(from).TrimEnd('\\');
            if (!Directory.Exists(to)) Directory.CreateDirectory(to);
            string dst = Path.GetFullPath(to).TrimEnd('\\');
            if (string.Equals(src, dst, StringComparison.OrdinalIgnoreCase) || dst.StartsWith(src + "\\", StringComparison.OrdinalIgnoreCase)) return res;
            foreach (string f in Directory.GetFiles(src, "*", SearchOption.AllDirectories))
            {
                try
                {
                    string rel = f.Substring(src.Length + 1);
                    string target = Path.Combine(dst, rel);
                    string dir = Path.GetDirectoryName(target);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    if (File.Exists(target))
                    {
                        if (new FileInfo(target).Length == new FileInfo(f).Length && Sha(target) == Sha(f))
                        {
                            if (!copy) File.Delete(f);
                            res.Same++; continue;
                        }
                        string bs = Path.GetFileNameWithoutExtension(f), ext = Path.GetExtension(f); int n = 2;
                        do { target = Path.Combine(dir, string.Format("{0} ({1}){2}", bs, n, ext)); n++; } while (File.Exists(target));
                        res.Renamed++;
                    }
                    if (copy) File.Copy(f, target); else File.Move(f, target);
                    res.Moved++;
                }
                catch (Exception e) { res.Failed++; App.Log("documentos: no se pudo pasar " + f + ": " + e.Message); }
            }
            if (copy)
            {
                App.Log(string.Format("documentos: copiados de {0} a {1}: {2} ({3} con otro nombre), {4} ya estaban, {5} sin copiar", src, dst, res.Moved, res.Renamed, res.Same, res.Failed));
                return res;
            }
            // carpetas que quedaron vacías (las más hondas primero)
            List<string> dirs = new List<string>(Directory.GetDirectories(src, "*", SearchOption.AllDirectories));
            dirs.Sort(delegate(string a, string b) { return b.Length.CompareTo(a.Length); });
            foreach (string d in dirs) try { if (Directory.GetFileSystemEntries(d).Length == 0) Directory.Delete(d); } catch { }
            try { if (Directory.GetFileSystemEntries(src).Length == 0) Directory.Delete(src); } catch { }
            App.Log(string.Format("documentos: de {0} a {1}: {2} pasados ({3} con otro nombre porque ya había uno distinto), {4} ya estaban, {5} sin pasar", src, dst, res.Moved, res.Renamed, res.Same, res.Failed));
            return res;
        }

        // mismo nombre, o uno contiene todas las palabras del otro («ANA» y «ANA RUIZ»): la misma persona con el nombre corregido
        public static bool SameNameKey(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            if (a == b) return true;
            List<string> wa = new List<string>(), wb = new List<string>();
            foreach (string w in a.Split(' ')) if (w.Length > 0) wa.Add(w);
            foreach (string w in b.Split(' ')) if (w.Length > 0) wb.Add(w);
            if (wa.Count > wb.Count) { List<string> x = wa; wa = wb; wb = x; }
            foreach (string w in wa) if (!PsH.ContainsI(wb, w)) return false;
            return true;
        }

        // carpetas «nº - nombre» cuyo nº ya es de otro cliente: pasan al nº del cliente con ese nombre, si hay uno solo
        public static int RepairFolders(Tramite t, StoreData s)
        {
            int n = 0;
            string root = Root();
            if (root == null || s == null) return 0;
            string tdir = Path.Combine(root, SafeName(t.Sheet));
            if (!Directory.Exists(tdir)) return 0;
            Dictionary<string, ClientRec> byId = new Dictionary<string, ClientRec>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, ClientRec> byName = new Dictionary<string, ClientRec>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> many = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ClientRec c in s.Clients)
            {
                byId[IdOf(c.Id, c.Row)] = c;
                if (c.Del != null) continue;
                string k = Hand.HandKey(SafeName(Ps.Str(Store.Get(c.F, "name"))));
                if (k.Length > 0) { if (byName.ContainsKey(k)) many.Add(k); else byName[k] = c; }
            }
            foreach (string d in Directory.GetDirectories(tdir))
            {
                string dn = Path.GetFileName(d);
                int i = dn.IndexOf(" - ", StringComparison.Ordinal); if (i <= 0) continue;
                string fid = dn.Substring(0, i), fk = Hand.HandKey(dn.Substring(i + 3));
                if (fk.Length == 0) continue;
                ClientRec own; byId.TryGetValue(fid, out own);
                if (own != null && SameNameKey(fk, Hand.HandKey(SafeName(Ps.Str(Store.Get(own.F, "name")))))) continue;   // es de su cliente
                ClientRec c;
                if (many.Contains(fk) || !byName.TryGetValue(fk, out c)) continue;    // nadie con ese nombre, o varios: no se toca
                string nid = IdOf(c.Id, c.Row);
                if (nid == fid) continue;
                string target = Path.Combine(tdir, nid + " - " + SafeName(Ps.Str(Store.Get(c.F, "name"))));
                try
                {
                    if (Directory.Exists(target)) Merge(d, target, false); else Directory.Move(d, target);
                    App.Log(string.Format("documentos: {0}: carpeta «{1}» pasa a «{2}» (el nº {3} es ahora de otro cliente)", t.Sheet, dn, Path.GetFileName(target), fid));
                    n++;
                }
                catch (Exception e) { App.Log("documentos: no se pudo renombrar «" + dn + "»: " + e.Message); }
            }
            if (n > 0) UpdateMap();
            return n;
        }
        public static void RepairAll()
        {
            foreach (Tramite t in Tramite.All)
            {
                try { StoreData s = Store.Read(t); if (s != null) RepairFolders(t, s); }
                catch (Exception e) { App.Log("documentos (" + t.Sheet + "): " + e.Message); }
            }
        }

        // nº de documentos por expediente del trámite actual (para la columna de la lista)
        public static void UpdateMap()
        {
            Map = PsH.New();
            try
            {
                string tdir = TramiteDir(G.T);
                if (tdir == null || !Directory.Exists(tdir)) return;
                foreach (string d in Directory.GetDirectories(tdir))
                {
                    string id = Path.GetFileName(d); int i = id.IndexOf(" - ", StringComparison.Ordinal); if (i > 0) id = id.Substring(0, i);
                    Map[id] = VisibleFiles(d).Count;
                }
            }
            catch { }
        }
        // Get-ChildItem -File sin -Force: sin los ocultos (desktop.ini…)
        public static List<FileInfo> VisibleFiles(string dir)
        {
            List<FileInfo> l = new List<FileInfo>();
            foreach (FileInfo f in new DirectoryInfo(dir).GetFiles()) if ((f.Attributes & FileAttributes.Hidden) == 0) l.Add(f);
            return l;
        }
        public static int Count(ClientRow r)
        {
            string id = IdOf(r.Id, r.Row);
            return Map.ContainsKey(id) ? (int)Map[id] : 0;
        }
    }
}

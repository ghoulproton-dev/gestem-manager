// Etapa 5a: copia externa (USB, OneDrive, Google Drive, red). Copiado de Get-CloudRoots, Find-CompanyCloud, Export-ActivityLog,
// Backup-External, Sync-DocsBackup y Sync-DataBackup de GESTEM.ps1 (y de los botones «copiar ahora» / «buscar OneDrive»).
// Modo prueba: App.CloudRootsOverride hace de «nube» (cs\data\nube) y App.Guard para cualquier escritura fuera de las carpetas de prueba.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Gestem
{
    public static class ExtBackup
    {
        // nube de la empresa: primero OneDrive de empresa (Microsoft 365), luego Google Drive, luego OneDrive personal
        public static List<string> CloudRoots()
        {
            List<string> c = new List<string>();
            if (App.CloudRootsOverride != null) c.AddRange(App.CloudRootsOverride);
            else
            {
                string e = Environment.GetEnvironmentVariable("OneDriveCommercial"); if (!string.IsNullOrEmpty(e)) c.Add(e);
                foreach (DriveInfo d in DriveInfo.GetDrives())
                {
                    try { if (d.IsReady) foreach (string n in new string[] { "Mi unidad", "My Drive", "Mon Drive" }) c.Add(Path.Combine(d.RootDirectory.FullName, n)); } catch { }
                }
                string up = Environment.GetEnvironmentVariable("USERPROFILE") ?? "";
                c.Add(Path.Combine(up, @"Google Drive\Mi unidad")); c.Add(Path.Combine(up, "Google Drive"));
                foreach (string v in new string[] { "OneDrive", "OneDriveConsumer" }) { e = Environment.GetEnvironmentVariable(v); if (!string.IsNullOrEmpty(e)) c.Add(e); }
            }
            List<string> o = new List<string>();
            foreach (string x in c) if (!string.IsNullOrEmpty(x) && (Directory.Exists(x) || File.Exists(x)) && !PsH.ContainsI(o, x)) o.Add(x);    // Select-Object -Unique (sin mayúsculas)
            return o;
        }
        // Get-ChildItem sin -Force: ni archivos ni carpetas ocultos (desktop.ini de las carpetas de documentos)
        static bool Shown(string p) { try { return (File.GetAttributes(p) & FileAttributes.Hidden) == 0; } catch { return false; } }
        static List<string> Files(string dir, string pattern, bool recurse)
        {
            List<string> l = new List<string>();
            try { foreach (string f in Directory.GetFiles(dir, pattern)) if (Shown(f)) l.Add(f); } catch { }
            if (recurse) { try { foreach (string d in Directory.GetDirectories(dir)) if (Shown(d)) l.AddRange(Files(d, pattern, true)); } catch { } }
            return l;
        }
        public static string FindCompanyCloud() { List<string> r = CloudRoots(); return r.Count > 0 ? Path.Combine(r[0], "GESTEM copias") : null; }

        // registro de actividad legible en la carpeta externa: un archivo por PC, lo más reciente arriba, SIN nombres de clientes
        static readonly Regex LogRe = new Regex(@"^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  \[([^\]]*)\] (.+)$", RegexOptions.IgnoreCase);
        public static void ExportActivityLog(string dir)
        {
            try
            {
                if (!File.Exists(App.LogPath)) return;
                string[] all = File.ReadAllLines(App.LogPath, Encoding.UTF8);
                List<string> lines = new List<string>(all.Length > 4000 ? new ArraySegment<string>(all, all.Length - 4000, 4000) : (IEnumerable<string>)all);
                lines.Reverse();
                List<string> outl = new List<string>();
                List<string> keys = Tx.AdmKeysLongestFirst();
                foreach (string ln in lines)
                {
                    Match m = LogRe.Match(ln); if (!m.Success) continue;
                    string d = m.Groups[3].Value + "/" + m.Groups[2].Value + "/" + m.Groups[1].Value + " " + m.Groups[4].Value, who = m.Groups[5].Value, txt = m.Groups[6].Value;
                    if (PsH.Is(txt, "^(Load-Data|Save-FromForm|clic |DEBUG|pestaña no leída|copia externa: |icono:)")) continue;
                    Match c = Regex.Match(txt, @"^(añadió|editó) cliente (\S+) · .* · (.+)$", RegexOptions.IgnoreCase);    // quita el nombre del cliente
                    if (c.Success) txt = c.Groups[1].Value + " cliente " + c.Groups[2].Value + " · " + c.Groups[3].Value;
                    foreach (string k in keys)
                    {
                        if (k.Length >= 8 && txt.StartsWith(k, StringComparison.CurrentCulture)) { string v = Tx.AdmEn(k); if (!string.IsNullOrEmpty(v)) txt = v + txt.Substring(k.Length); break; }
                    }
                    string label = string.Equals(who, "ADMIN", StringComparison.CurrentCultureIgnoreCase) ? "ADMIN" : who.Length > 0 ? who : "system";
                    outl.Add(string.Format("{0}   {1,-12} {2}", d, label, txt));
                    if (outl.Count >= 2000) break;
                }
                List<string> head = new List<string> { "GESTEM MANAGER - activity on PC " + Environment.MachineName, "Updated: " + App.Now().ToString("dd/MM/yyyy HH:mm") + "  ·  newest first  ·  no client names", "" };
                head.AddRange(outl);
                string file = Path.Combine(dir, string.Format("actividad GESTEM - {0}.txt", Environment.MachineName));
                App.Guard(file, "copia externa");
                File.WriteAllLines(file, head.ToArray(), new UTF8Encoding(true));
            }
            catch { }
        }

        // copia completa del Excel en la carpeta externa: un archivo por día, 60 días; también ajustes, documentos y datos
        public static bool BackupExternal(bool quiet, Action<string> showError)
        {
            string dir = Ps.Str(Config.Cur.ExtBackupDir).Trim();
            if (dir.Length == 0 || string.IsNullOrEmpty(Config.Cur.ExcelPath)) return false;
            try
            {
                App.Guard(dir, "copia externa");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string src = Config.Cur.ExcelPath;
                string name = Path.GetFileNameWithoutExtension(src), ext = Path.GetExtension(src);
                string dst = Path.Combine(dir, string.Format("{0}_{1:yyyy-MM-dd}{2}", name, App.Now(), ext));
                // lectura compartida: funciona aunque el Excel esté abierto
                using (FileStream fin = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (FileStream fout = new FileStream(dst + ".tmp", FileMode.Create, FileAccess.Write)) fin.CopyTo(fout);
                if (File.Exists(dst)) File.Delete(dst);
                File.Move(dst + ".tmp", dst);
                // también los ajustes del programa (empresa, logo, compañeras, idioma…) para poder restaurarlo todo en otro PC
                string setDir = Path.Combine(dir, "ajustes GESTEM");
                if (!Directory.Exists(setDir)) Directory.CreateDirectory(setDir);
                if (Directory.Exists(App.DataDir))
                    foreach (string f in Files(App.DataDir, "*", false))
                    {
                        string fn = Path.GetFileName(f);
                        if (string.Equals(fn, "config.json", StringComparison.CurrentCultureIgnoreCase) || fn.StartsWith("empresa.", StringComparison.CurrentCultureIgnoreCase)) File.Copy(f, Path.Combine(setDir, fn), true);
                    }
                List<string> old = Files(dir, name + "_????-??-??" + ext, false);
                old.Sort(delegate(string a, string b) { return -StringComparer.CurrentCultureIgnoreCase.Compare(Path.GetFileName(a), Path.GetFileName(b)); });
                for (int i = 60; i < old.Count; i++) { try { File.Delete(old[i]); } catch { } }
                Config.Cur.LastExtBackup = App.Now().ToString("yyyy-MM-dd HH:mm");
                Config.Save();
                App.Log("copia externa: " + dst);
                SyncDocsBackup(dir);
                SyncDataBackup(dir);
                ExportActivityLog(dir);
                return true;
            }
            catch (Exception e)
            {
                App.Log("copia externa fallida: " + e.Message);
                if (!quiet && showError != null) showError(Tx.L("extFail") + Environment.NewLine + e.Message);
                return false;
            }
        }

        // documentos escaneados: solo lo nuevo o cambiado
        public static void SyncDocsBackup(string dir)
        {
            string root = Docs.Root();
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return;
            string dstRoot = Path.Combine(dir, "documentos");
            foreach (string f in Files(root, "*", true))
            {
                try
                {
                    string rel = f.Substring(root.Length).TrimStart('\\');
                    string dst = Path.Combine(dstRoot, rel), di = Path.GetDirectoryName(dst);
                    App.Guard(dst, "copia externa");
                    if (!Directory.Exists(di)) Directory.CreateDirectory(di);
                    FileInfo fi = new FileInfo(f), o = new FileInfo(dst);
                    if (o.Exists && o.Length == fi.Length && o.LastWriteTimeUtc >= fi.LastWriteTimeUtc) continue;
                    File.Copy(f, dst, true);
                }
                catch { }
            }
        }

        // datos de GESTEM («GESTEM datos\*.json»)
        public static void SyncDataBackup(string dir)
        {
            string root = Store.Dir();
            if (string.IsNullOrEmpty(root) || !Directory.Exists(root)) return;
            string dst = Path.Combine(dir, "datos GESTEM");
            App.Guard(dst, "copia externa");
            if (!Directory.Exists(dst)) Directory.CreateDirectory(dst);
            foreach (string f in Files(root, "*.json", false)) { try { File.Copy(f, Path.Combine(dst, Path.GetFileName(f)), true); } catch { } }
        }
    }

    public partial class MainWin
    {
        void WireExt()
        {
            U<Button>("ExtOneDriveBtn").Click += delegate
            {
                string c = ExtBackup.FindCompanyCloud();
                if (c != null) { U<TextBox>("ExtDirBox").Text = c; UpdateExtBackupInfo(); } else ShowError(L("extNoCloud"));
            };
            U<Button>("ExtNowBtn").Click += delegate { ExtNow(); };
            // al cerrar (después del aviso de ficha a medias): registro y copia externa
            Win.Closing += delegate(object s, System.ComponentModel.CancelEventArgs e)
            {
                if (e.Cancel) return;
                if (App.User.Length > 0) App.Log("cerró GESTEM");
                if (Ps.Str(Config.Cur.ExtBackupDir).Length > 0 && !string.IsNullOrEmpty(Config.Cur.ExcelPath)) { try { ExtBackup.BackupExternal(true, null); } catch { } }
            };
        }

        public bool BackupExternal(bool quiet) { return ExtBackup.BackupExternal(quiet, ShowError); }

        public void ExtNow()
        {
            string old = Config.Cur.ExtBackupDir; Config.Cur.ExtBackupDir = Ps.Str(U<TextBox>("ExtDirBox").Text).Trim();
            Win.Cursor = Cursors.Wait;
            bool ok = BackupExternal(false);
            Win.Cursor = null;
            if (!ok) Config.Cur.ExtBackupDir = old; else { Config.Save(); SetStatus(L("extDone")); }
            UpdateExtBackupInfo();
        }

        // al abrir, después de cargar: primera vez sin carpeta externa → la nube de la empresa si este PC la tiene; luego la copia
        void StartupExtBackup()
        {
            if (Ps.Str(Config.Cur.ExtBackupDir).Length == 0)
            {
                string c = ExtBackup.FindCompanyCloud();
                if (c != null) { Config.Cur.ExtBackupDir = c; Config.Save(); App.Log("copia externa configurada automáticamente: " + c); }
            }
            if (Ps.Str(Config.Cur.ExtBackupDir).Length > 0) BackupExternal(true);
        }
    }
}

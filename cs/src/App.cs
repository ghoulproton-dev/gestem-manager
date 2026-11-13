// Rutas, registro (gestem.log), recursos y textos — igual que GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Gestem
{
    public static class App
    {
        public const string Name = "GESTEM MANAGER";
        public const string Version = "1.7.3";          // misma versión que la de PowerShell que copia
        public const string UpdateBase = "https://raw.githubusercontent.com/ghoulproton-dev/gestem-manager/main";

        public static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');
        public static string DataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "GESTEM");
        public static string CfgPath = Path.Combine(DataDir, "config.json");
        public static string BackupDir = Path.Combine(DataDir, "copias");
        public static string LogPath = Path.Combine(DataDir, "gestem.log");

        public static string User = "";
        public static bool IsAdmin = false;
        public static bool IsOwnerPC = false;

        public static void EnsureDirs()
        {
            Directory.CreateDirectory(DataDir);
            Directory.CreateDirectory(BackupDir);
        }

        // Write-Log
        public static void Log(string msg)
        {
            string who = (IsAdmin || (IsOwnerPC && User.Length > 0)) ? "ADMIN" : User;
            try
            {
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", Ps.Inv) + "  [" + who + "] " + msg + "\r\n";
                File.AppendAllText(LogPath, line, new UTF8Encoding(!File.Exists(LogPath)));
            }
            catch { }
        }

        // archivos de datos incluidos dentro del programa (data\*.json, main.xaml)
        static readonly Dictionary<string, object> resCache = new Dictionary<string, object>();
        public static string ResText(string name)
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (s == null) throw new FileNotFoundException("recurso interno no encontrado: " + name);
                using (StreamReader r = new StreamReader(s, Encoding.UTF8)) return r.ReadToEnd();
            }
        }
        public static object ResJson(string name)
        {
            object v;
            if (!resCache.TryGetValue(name, out v)) { v = Json.Parse(ResText(name)); resCache[name] = v; }
            return v;
        }
    }

    // textos en 4 idiomas (L 'clave')
    public static class Tx
    {
        static Dictionary<string, object> text;
        public static string Lang = "es";
        public static readonly string[] Langs = { "es", "en", "fr", "ar" };

        public static string L(string key)
        {
            if (text == null) text = Json.Obj(App.ResJson("text.json"));
            object lt; object v;
            if (text.TryGetValue(Lang, out lt) && Json.Obj(lt) != null && Json.Obj(lt).TryGetValue(key, out v)) return Ps.Str(v);
            if (text.TryGetValue("es", out lt) && Json.Obj(lt).TryGetValue(key, out v)) return Ps.Str(v);
            return null;
        }
        public static string F(string key, params object[] args)
        {
            string s = L(key) ?? "";
            object[] a = new object[args.Length];
            for (int i = 0; i < args.Length; i++) a[i] = (args[i] is string || args[i] == null) ? Ps.Str(args[i]) : args[i];
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, s, a);   // -f de PowerShell usa la cultura del PC
        }

        // AdmT: área de administrador (las claves son el texto en español)
        static Dictionary<string, object> adm;
        public static string AdmT(string es)
        {
            if (adm == null)
            {
                adm = new Dictionary<string, object>();
                Dictionary<string, object> all = Json.Obj(App.ResJson("admin_tr.json"));
                foreach (string part in new string[] { "adm", "help", "about" })
                    foreach (KeyValuePair<string, object> kv in Json.Obj(all[part])) adm[kv.Key] = kv.Value;
            }
            object tr;
            if (Lang != "es" && adm.TryGetValue(es, out tr)) { string v = Json.Str(tr, Lang); if (!string.IsNullOrEmpty(v)) return v; }
            return es;
        }
    }
}

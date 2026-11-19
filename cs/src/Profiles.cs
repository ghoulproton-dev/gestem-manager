// Perfiles de usuario (solo GESTEM.exe): contraseña e icono de cada persona. %APPDATA%\GESTEM\usuarios.json (y fotos\ para las fotos).
// La contraseña nunca se guarda: solo un hash PBKDF2-SHA256 con sal. Con Firebase encendido se sincronizan entre PCs SOLO el hash,
// la sal y el estado (gestorias/elmahi/usuarios); el icono y la foto se quedan en cada PC. Gana el último cambio («Updated», UTC).
// control.json «resetPass»: ["SAFAE"] restablece una vez por petición (ResetSeen, sincronizado: otro PC no lo repite); al quitar el
// nombre de la lista se puede volver a pedir.
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Gestem
{
    public class Profile
    {
        public string Name = "", Hash = "", Salt = "", Updated = "", Color = "", Photo = "";
        public int Iter;
        public bool ResetSeen;
        public bool HasPassword { get { return Hash.Length > 0; } }
        public Profile CopySync() { return new Profile { Name = Name, Hash = Hash, Salt = Salt, Updated = Updated, Iter = Iter, ResetSeen = ResetSeen }; }
    }

    public static class Profiles
    {
        public const int Iterations = 100000, MinLen = 4;
        public static readonly string[] Colors = { "#9E1820", "#C2410C", "#B45309", "#15803D", "#0F766E", "#1D4ED8", "#6D28D9", "#BE185D" };
        public static string FilePath { get { return Path.Combine(App.DataDir, "usuarios.json"); } }
        public static string PhotoDir { get { return Path.Combine(App.DataDir, "fotos"); } }
        public static Func<string> NowUtc = delegate { return DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", Ps.Inv); };

        static Dictionary<string, Profile> all;
        static string loadedFrom;
        static Dictionary<string, Profile> All
        {
            get
            {
                if (all != null && loadedFrom == FilePath) return all;
                all = new Dictionary<string, Profile>(StringComparer.OrdinalIgnoreCase); loadedFrom = FilePath;
                if (File.Exists(FilePath))
                    foreach (KeyValuePair<string, object> kv in Json.Obj(Json.Parse(File.ReadAllText(FilePath, Encoding.UTF8))))
                    {
                        object o = kv.Value; int it; int.TryParse(Ps.Str(Json.Get(o, "Iter")), out it);
                        all[kv.Key] = new Profile { Name = kv.Key, Hash = Ps.Str(Json.Get(o, "Hash")), Salt = Ps.Str(Json.Get(o, "Salt")), Updated = Ps.Str(Json.Get(o, "Updated")),
                            Iter = it, ResetSeen = Ps.Str(Json.Get(o, "ResetSeen")) == "True", Color = Ps.Str(Json.Get(o, "Color")), Photo = Ps.Str(Json.Get(o, "Photo")) };
                    }
                return all;
            }
        }
        public static void Reload() { all = null; }

        static void Save()
        {
            Dictionary<string, object> d = new Dictionary<string, object>();
            foreach (Profile p in All.Values.OrderBy(x => x.Name, StringComparer.Ordinal))
            {
                Dictionary<string, object> o = new Dictionary<string, object>();
                o["Hash"] = p.Hash; o["Salt"] = p.Salt; o["Iter"] = p.Iter; o["Updated"] = p.Updated; o["ResetSeen"] = p.ResetSeen; o["Color"] = p.Color; o["Photo"] = p.Photo;
                d[p.Name] = o;
            }
            App.Guard(FilePath, "usuarios");
            string tmp = FilePath + ".tmp";
            File.WriteAllText(tmp, Json.Write(d), new UTF8Encoding(false));
            if (File.Exists(FilePath)) File.Replace(tmp, FilePath, null); else File.Move(tmp, FilePath);
        }

        static string Key(string name) { return Ps.Str(name).Trim().ToUpper(); }
        public static List<string> Names() { return All.Keys.OrderBy(x => x, StringComparer.CurrentCulture).ToList(); }
        public static Profile Get(string name) { Profile p; return All.TryGetValue(Key(name), out p) ? p : null; }
        public static bool HasPassword(string name) { Profile p = Get(name); return p != null && p.HasPassword; }
        static Profile GetOrNew(string name)
        {
            Profile p = Get(name);
            if (p == null) { p = new Profile { Name = Key(name) }; All[p.Name] = p; }
            return p;
        }

        // color por defecto: siempre el mismo para el mismo nombre
        public static string ColorOf(string name)
        {
            Profile p = Get(name); if (p != null && p.Color.Length > 0) return p.Color;
            int h = 0; foreach (char c in Key(name)) h = (h * 31 + c) & 0x7FFFFFFF;
            return Colors[h % Colors.Length];
        }
        public static string Initials(string name)
        {
            string[] w = Key(name).Split(new char[] { ' ', '-', '.', '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (w.Length == 0) return "?";
            if (w.Length == 1) return w[0].Length > 1 ? w[0].Substring(0, 2) : w[0];
            return w[0].Substring(0, 1) + w[1].Substring(0, 1);
        }
        public static string PhotoPath(string name) { Profile p = Get(name); return p != null && p.Photo.Length > 0 && File.Exists(Path.Combine(PhotoDir, p.Photo)) ? Path.Combine(PhotoDir, p.Photo) : null; }

        static byte[] Derive(string pw, byte[] salt, int iter)
        {
            using (Rfc2898DeriveBytes k = new Rfc2898DeriveBytes(pw, salt, iter, HashAlgorithmName.SHA256)) return k.GetBytes(32);
        }
        // "" = bien; si no, el motivo (texto de la interfaz)
        public static string Validate(string pw, string pw2)
        {
            if (Ps.Str(pw).Length < MinLen) return "pwShort";
            if (pw != pw2) return "pwDiff";
            return "";
        }
        public static void SetPassword(string name, string pw)
        {
            if (Ps.Str(pw).Length < MinLen) throw new ArgumentException("contraseña demasiado corta");
            byte[] salt = new byte[16]; using (RNGCryptoServiceProvider r = new RNGCryptoServiceProvider()) r.GetBytes(salt);
            Profile p = GetOrNew(name);
            p.Salt = Convert.ToBase64String(salt); p.Iter = Iterations; p.Hash = Convert.ToBase64String(Derive(pw, salt, Iterations)); p.Updated = NowUtc();
            Save();
            App.Log("contraseña de " + p.Name + " creada");
        }
        public static bool Check(string name, string pw)
        {
            Profile p = Get(name);
            if (p == null || !p.HasPassword || pw == null) return false;
            byte[] want, got;
            try { want = Convert.FromBase64String(p.Hash); got = Derive(pw, Convert.FromBase64String(p.Salt), p.Iter > 0 ? p.Iter : Iterations); } catch { return false; }
            int diff = want.Length ^ got.Length;
            for (int i = 0; i < Math.Min(want.Length, got.Length); i++) diff |= want[i] ^ got[i];    // sin atajos: el mismo tiempo acierte o no
            return diff == 0;
        }
        // administrador o control.json: esa persona crea una contraseña nueva al entrar
        public static void Reset(string name, string why)
        {
            Profile p = GetOrNew(name);
            p.Hash = ""; p.Salt = ""; p.Iter = 0; p.Updated = NowUtc();
            Save();
            App.Log("contraseña de " + p.Name + " restablecida (" + why + ")");
        }
        public static void SetIcon(string name, string color, string photoFile)
        {
            Profile p = GetOrNew(name); p.Color = Ps.Str(color); p.Photo = Ps.Str(photoFile); Save();
        }

        // control.json «resetPass»: una vez por petición; al quitar el nombre de la lista, se puede pedir otra vez
        public static List<string> ApplyRemoteReset(IEnumerable<string> names)
        {
            HashSet<string> want = new HashSet<string>(names.Select(Key).Where(x => x.Length > 0), StringComparer.OrdinalIgnoreCase);
            List<string> done = new List<string>(); bool changed = false;
            foreach (string n in want)
            {
                Profile p = Get(n);
                if (p == null || p.ResetSeen) continue;
                p.Hash = ""; p.Salt = ""; p.Iter = 0; p.ResetSeen = true; p.Updated = NowUtc(); changed = true; done.Add(p.Name);
                App.Log("contraseña de " + p.Name + " restablecida (control.json)");
            }
            foreach (Profile p in All.Values) if (p.ResetSeen && !want.Contains(p.Name)) { p.ResetSeen = false; changed = true; }    // sin tocar Updated: no es un cambio para los otros PCs
            if (changed) Save();
            return done;
        }
        public static List<string> ResetNames(object remote)
        {
            object v = Json.Get(remote, "resetPass");
            List<string> l = new List<string>();
            IList arr = v as IList;
            if (arr != null) { foreach (object x in arr) l.Add(Ps.Str(x)); }
            else if (v != null) l.Add(Ps.Str(v));
            return l;
        }

        // ---------- sincronización (solo hash, sal y estado) ----------
        public static List<Profile> Snapshot() { return All.Values.Where(p => p.Updated.Length > 0).Select(p => p.CopySync()).ToList(); }

        // lo de otro PC más nuevo que lo de aquí entra; devuelve si cambió algo
        public static bool Merge(IEnumerable<Profile> remote)
        {
            bool changed = false;
            foreach (Profile r in remote)
            {
                if (Key(r.Name).Length == 0) continue;
                Profile l = Get(r.Name);
                if (l != null && string.CompareOrdinal(r.Updated, l.Updated) <= 0) continue;
                l = GetOrNew(r.Name);
                l.Hash = r.Hash; l.Salt = r.Salt; l.Iter = r.Iter; l.Updated = r.Updated; l.ResetSeen = r.ResetSeen; changed = true;
                App.Log("usuario " + l.Name + ": " + (l.HasPassword ? "contraseña" : "restablecimiento") + " recibido de otro PC");
            }
            if (changed) Save();
            return changed;
        }

        public static Func<List<Profile>, List<Profile>> RemoteHook;    // pruebas: Firebase de mentira (recibe lo que se envía, devuelve lo que hay)

        // en segundo plano: envía lo de aquí más nuevo que lo de Firebase y devuelve lo que hay en Firebase (para Merge en la ventana)
        public static List<Profile> SyncOnce(string key, ref FbAuth auth, List<Profile> local)
        {
            if (RemoteHook != null) return RemoteHook(local);
            FbAuth a = auth != null && DateTime.Now < auth.Until ? auth : Fb.SignIn(key); auth = a;
            List<Profile> remote = Fetch(a);
            Dictionary<string, Profile> byName = new Dictionary<string, Profile>(StringComparer.OrdinalIgnoreCase);
            foreach (Profile r in remote) byName[r.Name] = r;
            List<Profile> push = local.Where(l => { Profile r; return !byName.TryGetValue(l.Name, out r) || string.CompareOrdinal(l.Updated, r.Updated) > 0; }).ToList();
            if (push.Count > 0) Push(a, push);
            foreach (Profile p in push) byName[p.Name] = p;
            return byName.Values.ToList();
        }

        static string DocName(string name) { return Fb.Base + "/usuarios/u-" + BitConverter.ToString(Encoding.UTF8.GetBytes(Key(name))).Replace("-", ""); }

        static List<Profile> Fetch(FbAuth a)
        {
            Fb.CheckName(Fb.Base);
            object q = Fb.D("structuredQuery", Fb.D("from", new List<object> { Fb.D("collectionId", "usuarios") }));
            object res = Fb.J("POST", "https://firestore.googleapis.com/v1/" + Fb.Base + ":runQuery", q, a.Id, "runQuery " + Fb.Base + "/usuarios");
            List<Profile> l = new List<Profile>(); IList arr = res as IList;
            if (arr != null) foreach (object x in arr)
            {
                object doc = Json.Get(x, "document"); if (doc == null) continue;
                object f = Json.Get(doc, "fields"); int it; int.TryParse(Fb.S(f, "iter", "integerValue"), out it);
                l.Add(new Profile { Name = Fb.S(f, "name", "stringValue"), Hash = Fb.S(f, "hash", "stringValue"), Salt = Fb.S(f, "salt", "stringValue"), Iter = it,
                    Updated = Fb.S(f, "updated", "stringValue"), ResetSeen = Fb.S(f, "resetSeen", "booleanValue") == "True" });
            }
            return l;
        }

        static void Push(FbAuth a, List<Profile> ps)
        {
            List<object> writes = new List<object>(); List<string> names = new List<string>();
            foreach (Profile p in ps)
            {
                string n = DocName(p.Name); names.Add(n);
                writes.Add(Fb.D("update", Fb.D("name", n, "fields", Fb.D("name", Fb.D("stringValue", p.Name), "hash", Fb.D("stringValue", p.Hash), "salt", Fb.D("stringValue", p.Salt),
                        "iter", Fb.D("integerValue", p.Iter.ToString(Ps.Inv)), "updated", Fb.D("stringValue", p.Updated), "resetSeen", Fb.D("booleanValue", p.ResetSeen),
                        "pc", Fb.D("stringValue", Environment.MachineName))),
                    "updateTransforms", new List<object> { Fb.D("fieldPath", "mod", "setToServerValue", "REQUEST_TIME") }));
            }
            Fb.Commit(a, writes, names);
        }
    }
}

// Firebase: estado de este PC (%APPDATA%\GESTEM\firebase.json). Copiado de Get-FbState / Save-FbState.
// La sincronización en sí (Step-FbSync, Join-FbDocs…) es de la etapa 5b.
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Gestem
{
    public class PendItem { public string T = "", H = ""; public bool N; }

    public class FbState
    {
        public static string Path = System.IO.Path.Combine(App.DataDir, "firebase.json");
        public static bool Applying;        // $script:FbApplying: lo que llega de Firebase no se vuelve a enviar

        // Joined: '' = aún no se ha unido (primero se recibe todo), 'push' = unido, falta enviar lo suyo, 'yes' = normal
        public string LastMod = "", Joined = "";
        public Hashtable Pend = PsH.New();     // uid -> PendItem (mismo orden que @{} de PowerShell)

        public static FbState Load()
        {
            FbState s = new FbState();
            if (!File.Exists(Path)) return s;
            // si no se puede leer, mejor fallar que perder lo pendiente de enviar
            object j = Json.Parse(File.ReadAllText(Path, Encoding.UTF8));
            s.LastMod = Ps.Str(Json.Get(j, "lastMod")); s.Joined = Ps.Str(Json.Get(j, "joined"));
            Dictionary<string, object> p = Json.Obj(Json.Get(j, "pend"));
            if (p != null) foreach (KeyValuePair<string, object> kv in p)
            {
                PendItem i = new PendItem(); i.T = Ps.Str(Json.Get(kv.Value, "t")); i.H = Ps.Str(Json.Get(kv.Value, "h"));
                object n = Json.Get(kv.Value, "n"); i.N = n is bool && (bool)n;
                s.Pend[kv.Key] = i;
            }
            return s;
        }

        public void Save()
        {
            Dictionary<string, object> pend = new Dictionary<string, object>();
            foreach (DictionaryEntry e in Pend)
            {
                PendItem i = (PendItem)e.Value;
                Dictionary<string, object> x = new Dictionary<string, object>(); x["t"] = i.T; x["h"] = i.H; x["n"] = i.N;
                pend[(string)e.Key] = x;
            }
            Dictionary<string, object> o = new Dictionary<string, object>(); o["lastMod"] = LastMod; o["joined"] = Joined; o["pend"] = pend;
            App.Guard(Path, "firebase.json");
            string tmp = Path + ".tmp";
            File.WriteAllText(tmp, Json.Write(o), new UTF8Encoding(true));
            if (File.Exists(Path)) File.Delete(Path);
            File.Move(tmp, Path);
        }
    }
}

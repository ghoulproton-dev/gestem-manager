// Etapa 5d: Firebase (clientes compartidos entre PCs). Copiado de Connect-Firebase, $script:FbSyncScript, Step-FbSync, Test-FbSamePerson,
// Get-FbWantRow, Join-FbDocs, Get-FbRow, Receive-FbDocs, Get-StoreFields -Fb, Set-FbSyncSwitch e Invoke-LiveRefresh de GESTEM.ps1.
// Seguridad: todo pasa por Fb.Doc / Fb.CheckName. Sin --real solo se puede usar gestorias/elmahi/_tests/... (nunca los clientes de
// verdad) y la sincronización no corre nunca (modo prueba); los campos Local (contraseñas) nunca se suben (Store.FbFields).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Gestem
{
    public class FbAuth { public string Id = "", Refresh = "", Uid = ""; public DateTime Until; }
    public class FbDoc { public string Uid = "", T = "", Id = "", Datos = "", Updated = "", By = "", Pc = ""; public bool Deleted; public object O; }
    public class FbPush { public string Uid = "", T = "", Id = "", Datos = "", Updated = "", By = "", H = ""; public bool N, Deleted; }
    public class FbResult { public FbAuth Auth; public List<FbDoc> Pulled = new List<FbDoc>(); public List<FbPush> Pushed = new List<FbPush>(); public string Mod = "", Err = ""; }

    public static class Fb
    {
        public const string Project = "gestemgestem", ApiKey = "AIzaSyBHdQjaQUCezvGtWB9G_KXaMigp8tj89Mo", Email = "elmahi@gestem.app";
        public static string Gest = "elmahi";            // pruebas: «elmahi/_tests/...»
        public static FbAuth Auth;                       // $script:FbAuth
        public static string LastErr = "";               // $script:FbLastErr
        public static Action<string> Notify;             // Show-Toast (fbClash / fbRenum)
        public static Action<string, string> Trace;      // pruebas: cada ruta de Firestore que se toca (método, ruta)
        public static string Root { get { return "projects/" + Project + "/databases/(default)/documents"; } }
        public static string Base { get { return Root + "/gestorias/" + Gest; } }

        // la única puerta a Firestore: sin --real, solo dentro de gestorias/elmahi/_tests/
        public static void CheckName(string name)
        {
            bool test = Gest.StartsWith("elmahi/_tests/", StringComparison.Ordinal) && Gest.Length > 14 && !Gest.Contains("..") && !Gest.EndsWith("/");
            if (!test && !(App.Real && Gest == "elmahi")) throw new InvalidOperationException("Firebase: gestoría no permitida sin --real: " + Gest);
            if (!(name + "/").StartsWith(Base + "/", StringComparison.Ordinal) && name != Base) throw new InvalidOperationException("Firebase: ruta fuera de gestorias/" + Gest + ": " + name);
        }

        internal static object J(string method, string uri, object body, string token, string tracePath)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            if (Trace != null && tracePath != null) Trace(method, tracePath);
            using (HttpClient hc = new HttpClient())
            {
                hc.Timeout = TimeSpan.FromSeconds(20);
                HttpRequestMessage rq = new HttpRequestMessage(new HttpMethod(method), uri);
                if (token != null) rq.Headers.Add("Authorization", "Bearer " + token);
                if (body != null) rq.Content = new StringContent(Json.WriteCompact(body), Encoding.UTF8, "application/json");
                using (HttpResponseMessage r = hc.SendAsync(rq).Result)
                {
                    string txt = r.Content.ReadAsStringAsync().Result;
                    if (!r.IsSuccessStatusCode) throw new Exception("(" + (int)r.StatusCode + ") " + r.ReasonPhrase + " · " + Regex.Replace(txt, @"\s+", " "));
                    return txt.Trim().Length > 0 ? Json.Parse(txt) : null;
                }
            }
        }
        internal static Dictionary<string, object> D(params object[] kv) { Dictionary<string, object> d = new Dictionary<string, object>(); for (int i = 0; i < kv.Length; i += 2) d[(string)kv[i]] = kv[i + 1]; return d; }
        internal static string S(object fields, string k, string type) { return Ps.Str(Json.Get(Json.Get(fields, k), type)); }

        public static FbAuth SignIn(string key)
        {
            object r = J("POST", "https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=" + ApiKey, D("email", Email, "password", key, "returnSecureToken", true), null, null);
            int exp; int.TryParse(Ps.Str(Json.Get(r, "expiresIn")), out exp);
            return new FbAuth { Id = Ps.Str(Json.Get(r, "idToken")), Refresh = Ps.Str(Json.Get(r, "refreshToken")), Uid = Ps.Str(Json.Get(r, "localId")), Until = DateTime.Now.AddSeconds(exp - 120) };
        }

        // $script:FbSyncScript: entrar, recibir lo cambiado desde lastMod, enviar lo pendiente (en segundo plano, sin tocar la ventana)
        public static FbResult SyncOnce(string key, FbAuth auth, string lastMod, List<FbPush> push, string pc)
        {
            FbResult o = new FbResult(); o.Auth = auth;
            try
            {
                CheckName(Base);
                FbAuth a = auth;
                if (a == null || DateTime.Now >= a.Until) { a = SignIn(key); o.Auth = a; }
                Dictionary<string, object> q = D("from", new List<object> { D("collectionId", "clientes") }, "orderBy", new List<object> { D("field", D("fieldPath", "mod"), "direction", "ASCENDING") });
                if (!string.IsNullOrEmpty(lastMod)) q["where"] = D("fieldFilter", D("field", D("fieldPath", "mod"), "op", "GREATER_THAN", "value", D("timestampValue", lastMod)));
                object res = J("POST", "https://firestore.googleapis.com/v1/" + Base + ":runQuery", D("structuredQuery", q), a.Id, "runQuery " + Base + "/clientes");
                IList arr = res as IList;
                if (arr != null) foreach (object x in arr)
                {
                    object doc = Json.Get(x, "document"); if (doc == null) continue;
                    object f = Json.Get(doc, "fields"); string name = Ps.Str(Json.Get(doc, "name"));
                    object del = Json.Get(Json.Get(f, "deleted"), "booleanValue");
                    o.Pulled.Add(new FbDoc { Uid = name.Substring(name.LastIndexOf('/') + 1), T = S(f, "t", "stringValue"), Id = S(f, "id", "stringValue"), Datos = S(f, "datos", "stringValue"),
                        Updated = S(f, "updated", "stringValue"), By = S(f, "by", "stringValue"), Pc = S(f, "pc", "stringValue"), Deleted = del is bool && (bool)del });
                    o.Mod = S(f, "mod", "timestampValue");
                }
                // no se envía lo que otro PC cambió después (gana el último cambio) ni un cliente nuevo cuyo nº ya trae otro PC (se renumera antes)
                Hashtable newer = PsH.New(), ids = PsH.New();
                foreach (FbDoc d in o.Pulled) if (!string.Equals(d.Pc, pc, StringComparison.CurrentCultureIgnoreCase)) { newer[d.Uid] = d.Updated; if (d.Id.Length > 0) ids[d.T + "|" + d.Id] = d.Uid; }
                List<FbPush> todo = new List<FbPush>();
                foreach (FbPush p in push)
                {
                    if (newer.ContainsKey(p.Uid) && Gt((string)newer[p.Uid], p.Updated)) continue;
                    if (p.Id.Length > 0 && ids.ContainsKey(p.T + "|" + p.Id) && !string.Equals((string)ids[p.T + "|" + p.Id], p.Uid, StringComparison.CurrentCultureIgnoreCase)) continue;
                    todo.Add(p);
                }
                for (int i = 0; i < todo.Count; i += 400)
                {
                    List<FbPush> chunk = todo.GetRange(i, Math.Min(400, todo.Count - i));
                    List<object> w = new List<object>();
                    foreach (FbPush d in chunk)
                    {
                        string name = Base + "/clientes/" + d.Uid; CheckName(name);
                        Dictionary<string, object> fields = D("t", D("stringValue", d.T), "id", D("stringValue", d.Id), "datos", D("stringValue", d.Datos), "updated", D("stringValue", d.Updated),
                            "by", D("stringValue", d.By), "pc", D("stringValue", pc), "deleted", D("booleanValue", d.Deleted));
                        Dictionary<string, object> x = D("update", D("name", name, "fields", fields), "updateTransforms", new List<object> { D("fieldPath", "mod", "setToServerValue", "REQUEST_TIME") });
                        // cliente nuevo: si otro PC ya subió uno con ese id entretanto, falla todo y este PC se vuelve a unir (recibe primero)
                        if (d.N) x["currentDocument"] = D("exists", false);
                        w.Add(x);
                        if (Trace != null) Trace("WRITE", name);
                    }
                    J("POST", "https://firestore.googleapis.com/v1/" + Root + ":commit", D("writes", w), a.Id, null);
                    o.Pushed.AddRange(chunk);
                }
            }
            catch (Exception e)
            {
                Exception x = e is AggregateException && e.InnerException != null ? e.InnerException : e;
                o.Err = Regex.Replace(x.Message, @"\s+", " ");
            }
            return o;
        }
        static bool Gt(string a, string b) { return string.Compare(a, b, StringComparison.CurrentCultureIgnoreCase) > 0; }    // -gt de PowerShell con textos

        // pruebas: todos los documentos debajo de la gestoría de prueba, y borrarlos (nunca fuera de gestorias/elmahi/_tests/)
        public static List<string> ListAll(FbAuth a)
        {
            CheckName(Base);
            object res = J("POST", "https://firestore.googleapis.com/v1/" + Base + ":runQuery", D("structuredQuery", D("from", new List<object> { D("allDescendants", true) })), a.Id, "runQuery " + Base + "/**");
            List<string> l = new List<string>();
            IList arr = res as IList;
            if (arr != null) foreach (object x in arr) { object doc = Json.Get(x, "document"); if (doc != null) l.Add(Ps.Str(Json.Get(doc, "name"))); }
            return l;
        }
        // pruebas: todos los documentos debajo de gestorias/elmahi/_tests (listar la colección no lo dejan las reglas). El servidor solo
        // devuelve nombres entre _tests/! y _tests0/!; aun así, si llegara uno de fuera, error (no se borra nada)
        public static List<string> ListUnderTests(FbAuth a)
        {
            string p = Root + "/gestorias/elmahi/_tests";
            Func<string, string, object> ff = delegate(string op, string v) { return D("fieldFilter", D("field", D("fieldPath", "__name__"), "op", op, "value", D("referenceValue", v))); };
            object q = D("structuredQuery", D("from", new List<object> { D("allDescendants", true) }, "where", D("compositeFilter", D("op", "AND", "filters", new List<object> { ff("GREATER_THAN_OR_EQUAL", p + "/!"), ff("LESS_THAN", p + "0/!") }))));
            object res = J("POST", "https://firestore.googleapis.com/v1/" + Root + "/gestorias/elmahi:runQuery", q, a.Id, "runQuery gestorias/elmahi/_tests/**");
            List<string> l = new List<string>(); IList arr = res as IList;
            if (arr != null) foreach (object x in arr)
            {
                object doc = Json.Get(x, "document"); if (doc == null) continue;
                string n = Ps.Str(Json.Get(doc, "name"));
                if (!n.StartsWith(p + "/", StringComparison.Ordinal)) throw new InvalidOperationException("Firebase: la consulta de pruebas devolvió algo fuera de _tests: " + n);
                l.Add(n);
            }
            return l;
        }        public static void DeleteDoc(FbAuth a, string name)
        {
            CheckName(name);
            if (!Gest.StartsWith("elmahi/_tests/", StringComparison.Ordinal)) throw new InvalidOperationException("Firebase: borrar solo en pruebas");
            J("DELETE", "https://firestore.googleapis.com/v1/" + name, null, a.Id, name);
        }
        public static void Commit(FbAuth a, List<object> writes, List<string> names)
        {
            foreach (string n in names) { CheckName(n); if (Trace != null) Trace("WRITE", n); }
            J("POST", "https://firestore.googleapis.com/v1/" + Root + ":commit", D("writes", writes), a.Id, null);
        }

        // Get-StoreFields -Fb: lo que se sube (sin contraseñas)
        public static Dictionary<string, object> StoreFields(Tramite t, ClientRec c)
        {
            Dictionary<string, object> o = new Dictionary<string, object>();
            foreach (Field f in Store.FbFields(t)) o[f.Key] = Store.ToStoreValue(f, Store.Get(c.F, f.Key));
            return o;
        }

        // mismo cliente si coincide el NIE/pasaporte; si alguno no lo tiene, el nombre (sin tildes, mayúsculas ni espacios de más)
        public static bool SamePerson(Dictionary<string, object> F, object o)
        {
            string a = PsH.Rep(Xlsx.SheetKey(Ps.Str(Store.Get(F, "doc"))), "[^A-Z0-9]", ""), b = PsH.Rep(Xlsx.SheetKey(Ps.Str(Json.Get(o, "doc"))), "[^A-Z0-9]", "");
            if (a.Length > 0 && b.Length > 0) return string.Equals(a, b, StringComparison.CurrentCultureIgnoreCase);
            return string.Equals(Xlsx.SheetKey(Ps.Str(Store.Get(F, "name"))), Xlsx.SheetKey(Ps.Str(Json.Get(o, "name"))), StringComparison.CurrentCultureIgnoreCase);
        }

        // fila que corresponde a un nº de expediente (PREFIJO + número)
        public static int WantRow(Tramite t, Meta meta, string id)
        {
            if (string.IsNullOrEmpty(meta.IdPrefix)) return 0;
            Match m = Regex.Match(Ps.Str(id), "^" + Regex.Escape(meta.IdPrefix) + @"(\d+)$", RegexOptions.IgnoreCase);
            long n; if (!m.Success || !long.TryParse(m.Groups[1].Value, out n)) return 0;
            long w = t.FirstRow + n - 1;
            return w <= t.LastRow ? (int)w : 0;
        }

        static List<object> CloudVals(Tramite t, object o) { List<object> v = new List<object>(); foreach (Field f in Store.FbFields(t)) v.Add(Json.Get(o, f.Key)); return v; }
        static int FirstFree(Tramite t, HashSet<int> used) { for (int r = t.FirstRow; r <= t.LastRow; r++) if (!used.Contains(r)) return r; return 0; }

        // primera sincronización de este PC: ya se recibió todo; se compara con lo de este PC antes de enviar nada
        public static void Join(List<FbDoc> docs, string mod)
        {
            FbState st = FbState.Load(); st.Pend.Clear();
            foreach (Tramite t in Tramite.All)
            {
                List<FbDoc> cloud = new List<FbDoc>(); foreach (FbDoc d in docs) if (string.Equals(d.T, t.Id, StringComparison.CurrentCultureIgnoreCase)) cloud.Add(d);
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    if (s == null) { if (cloud.Count > 0) App.Log(string.Format("firebase: sin datos de {0}, {1} cliente(s) no aplicados", t.Sheet, cloud.Count)); continue; }
                    Hashtable byUid = PsH.New(), byId = PsH.New(); HashSet<int> want = new HashSet<int>();
                    foreach (FbDoc d in cloud)
                    {
                        d.O = Json.Parse(d.Datos); byUid[d.Uid] = d;
                        if (d.Id.Length > 0) byId[d.Id] = d;
                        int w = WantRow(t, s.Meta, d.Id); if (w > 0) want.Add(w);
                    }
                    // clientes de pestañas hechas a mano: una sola vez, al unirse, pasan a ser de GESTEM (nº libre que no use Firebase) y se suben
                    List<ClientRec> hand = new List<ClientRec>(); foreach (ClientRec c in s.Clients) if (c.Row >= Hand.RowBase) hand.Add(c);
                    Ps.Sort(hand, delegate(ClientRec a, ClientRec b) { return a.Row.CompareTo(b.Row); });
                    if (hand.Count > 0)
                    {
                        HashSet<int> used = new HashSet<int>(want); foreach (ClientRec x in s.Clients) if (x.Row < Hand.RowBase) used.Add(x.Row);
                        HashSet<string> uids = new HashSet<string>(); int n = 0;
                        foreach (ClientRec c in hand)
                        {
                            int row = FirstFree(t, used);
                            if (row == 0) { App.Log(string.Format("firebase: {0} lleno, quedan clientes de la pestaña a mano sin subir", t.Sheet)); break; }
                            string key = Hand.HandKey(Ps.Str(Store.Get(c.F, "name")));
                            string uid = Data.HandDelUid(t, key, "mano-"); int k2 = 2; while (!uids.Add(uid)) { uid = Data.HandDelUid(t, key, "mano-") + "-" + k2; k2++; }
                            used.Add(row);
                            c.Row = row; c.Id = Store.RowId(t, s.Meta, row); c.Uid = uid; c.Sync = true; c.New = false; c.H = null;
                            List<string> ad = new List<string>(); foreach (string x in s.Adopted) ad.Add(x); ad.Add(Ps.Str(Store.Get(c.F, "name")));
                            List<string> u2 = new List<string>(); foreach (string x in ad) if (Ps.Str(x).Trim().Length > 0 && !u2.Contains(x)) u2.Add(x);    // Select-Object -Unique (distingue mayúsculas)
                            s.Adopted = u2; n++;
                        }
                        if (n > 0) App.Log(string.Format("firebase: {0}: {1} cliente(s) de la pestaña hecha a mano pasan a GESTEM", t.Sheet, n));
                    }
                    HashSet<string> taken = new HashSet<string>(); List<ClientRec> clash = new List<ClientRec>();
                    foreach (ClientRec c in new List<ClientRec>(s.Clients))
                    {
                        if (c.Row >= Hand.RowBase) continue;
                        FbDoc d = c.Uid != null ? byUid[c.Uid] as FbDoc : null; if (d == null && !string.IsNullOrEmpty(c.Id)) d = byId[c.Id] as FbDoc;
                        if (d == null) { st.Pend[c.Uid ?? ""] = new PendItem { T = t.Id, H = "nuevo", N = true }; continue; }    // solo en este PC: se sube
                        if (SamePerson(c.F, d.O) && taken.Add(d.Uid))
                        {
                            c.Uid = d.Uid;
                            // distinto en los dos sitios: gana el cambio más reciente (si es el de Firebase, lo aplica Receive abajo)
                            if (!string.Equals(c.H, Store.FbHash(Store.FbKey(d.Id, d.Deleted), CloudVals(t, d.O)), StringComparison.CurrentCultureIgnoreCase) && Gt(Ps.Str(c.Updated), d.Updated))
                                st.Pend[c.Uid] = new PendItem { T = t.Id, H = "mas reciente", N = false };
                            continue;
                        }
                        clash.Add(c);
                    }
                    foreach (HandDelRec e in s.HandDel) if (e != null) st.Pend[e.Uid] = new PendItem { T = t.Id, H = "mano", N = false };    // eliminados de pestañas a mano: se suben
                    foreach (ClientRec c in clash)
                    {
                        // mismo nº que otro cliente de Firebase: el de este PC pasa al siguiente nº libre (que no sea de nadie en Firebase)
                        HashSet<int> used = new HashSet<int>(want); foreach (ClientRec x in s.Clients) if (!ReferenceEquals(x, c)) used.Add(x.Row);
                        int row = FirstFree(t, used);
                        if (row == 0) throw new GestemException(Tx.L("errFull"));
                        string old = c.Id;
                        c.Uid = Guid.NewGuid().ToString("N"); c.Row = row; c.Id = Store.RowId(t, s.Meta, row); c.Sync = false;
                        st.Pend[c.Uid] = new PendItem { T = t.Id, H = "renumerado", N = true };
                        App.Log(string.Format("firebase: nº {0} es otro cliente en Firebase; {1} de este PC pasa a {2}", old, Ps.Str(Store.Get(c.F, "name")), c.Id));
                        try { if (Notify != null) Notify(Tx.F("fbClash", old, c.Id, Ps.Str(Store.Get(c.F, "name")))); } catch { }
                    }
                    FbState.Applying = true; Store.Write(t, s);
                    try { Docs.RepairFolders(t, s); } catch (Exception e) { App.Log("documentos: " + e.Message); }    // carpetas de los renumerados
                }
                finally { FbState.Applying = false; Store.ExitLock(lk); }
            }
            st.Save();
            Receive(docs, true);    // lo de Firebase que falta o es más reciente
            st = FbState.Load(); st.Joined = "push"; if (!string.IsNullOrEmpty(mod)) st.LastMod = mod; st.Save();
            App.Log(string.Format("firebase: unión, recibidos {0} y por subir {1} cliente(s)", docs.Count, st.Pend.Count));
        }

        // fila para un cliente que llega de otro PC: la de su nº si está libre
        static int Row(Tramite t, StoreData s, ClientRec c, string id, FbState st)
        {
            Dictionary<int, ClientRec> byRow = new Dictionary<int, ClientRec>(); foreach (ClientRec x in s.Clients) if (!ReferenceEquals(x, c)) byRow[x.Row] = x;
            int free = 0; for (int r = t.FirstRow; r <= t.LastRow; r++) if (!byRow.ContainsKey(r)) { free = r; break; }
            int want = WantRow(t, s.Meta, id);
            if (want == 0) return c != null ? c.Row : free;
            ClientRec y; if (!byRow.TryGetValue(want, out y)) return want;
            if (string.Equals(y.Id, id, StringComparison.CurrentCultureIgnoreCase) && !string.IsNullOrEmpty(y.Uid) && st.Pend.ContainsKey(y.Uid))
            {
                // dos clientes nuevos con el mismo nº creados sin internet: el de este PC (el segundo en llegar) pasa al siguiente nº libre
                if (free == 0) throw new GestemException(Tx.L("errFull"));
                y.Row = free; y.Id = Store.RowId(t, s.Meta, free);
                ((PendItem)st.Pend[y.Uid]).H = "renumerado";
                App.Log(string.Format("firebase: nº {0} repetido con otro PC, el cliente de este PC pasa a {1}", id, y.Id));
                try { if (Notify != null) Notify(Tx.F("fbRenum", id, y.Id)); } catch { }
                return want;
            }
            if (c != null) return c.Row;
            App.Log(string.Format("firebase: nº {0} ya lo tiene otro cliente, el que llega va a la fila {1}", id, free == 0 ? "" : free.ToString()));
            return free;
        }

        // aplica lo que cambiaron otros PC; si este PC también lo cambió sin enviarlo todavía, gana el cambio más reciente
        public static void Receive(List<FbDoc> docs, bool all)
        {
            FbState st = FbState.Load();
            List<string> order = new List<string>(); Dictionary<string, List<FbDoc>> groups = new Dictionary<string, List<FbDoc>>(StringComparer.CurrentCultureIgnoreCase);    // Group-Object
            foreach (FbDoc d in docs)
            {
                if (!all && string.Equals(d.Pc, Environment.MachineName, StringComparison.CurrentCultureIgnoreCase)) continue;
                if (!groups.ContainsKey(d.T)) { groups[d.T] = new List<FbDoc>(); order.Add(d.T); }
                groups[d.T].Add(d);
            }
            foreach (string tid in order)
            {
                Tramite t = null; foreach (Tramite x in Tramite.All) if (string.Equals(x.Id, tid, StringComparison.CurrentCultureIgnoreCase)) { t = x; break; }
                if (t == null) { App.Log("firebase: trámite desconocido " + tid); continue; }
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    // shortcut: sin datos de este trámite en este PC no se aplica (como PowerShell); arreglar si aparece en el registro
                    if (s == null) { App.Log("firebase: sin datos de " + t.Sheet + ", cambios no aplicados"); continue; }
                    bool dirty = false;
                    foreach (FbDoc d in groups[tid])
                    {
                        object o = Json.Parse(d.Datos);
                        if (d.Uid.StartsWith("hand-", StringComparison.CurrentCultureIgnoreCase))
                        {
                            // cliente de una pestaña hecha a mano eliminado (o vuelto a poner) en otro PC
                            HandDelRec e = null; foreach (HandDelRec x in s.HandDel) if (x != null && string.Equals(x.Uid, d.Uid, StringComparison.CurrentCultureIgnoreCase)) { e = x; break; }
                            if (e != null && e.On == d.Deleted) continue;
                            if (e != null && string.Compare(Ps.Str(e.At), d.Updated, StringComparison.CurrentCultureIgnoreCase) >= 0) continue;    // el cambio de este PC es más reciente
                            st.Pend.Remove(d.Uid);
                            if (e == null) { e = new HandDelRec { Uid = d.Uid, Key = Hand.HandKey(Ps.Str(Json.Get(o, "name"))), Name = Ps.Str(Json.Get(o, "name")) }; s.HandDel.Add(e); }
                            e.On = d.Deleted; e.By = d.By; e.At = d.Updated;
                            if (e.On) s.Clients.RemoveAll(delegate(ClientRec x) { return x.Row >= Hand.RowBase && string.Equals(Hand.HandKey(Ps.Str(Store.Get(x.F, "name"))), e.Key, StringComparison.CurrentCultureIgnoreCase); });
                            else Hand.HandStamp = 0;    // vuelve: se vuelve a leer la pestaña
                            dirty = true;
                            App.Log(string.Format("firebase: {0} {1} (pestaña a mano) de {2}", e.On ? "eliminado" : "recuperado", e.Name, d.Pc));
                            continue;
                        }
                        ClientRec c = null; foreach (ClientRec x in s.Clients) if (string.Equals(x.Uid, d.Uid, StringComparison.CurrentCultureIgnoreCase)) { c = x; break; }
                        if (c != null && string.Equals(c.H, Store.FbHash(Store.FbKey(d.Id, d.Deleted), CloudVals(t, o)), StringComparison.CurrentCultureIgnoreCase)) continue;    // ya está igual
                        if (c != null && st.Pend.ContainsKey(d.Uid))
                        {
                            if (string.Compare(d.Updated, Ps.Str(c.Updated), StringComparison.CurrentCultureIgnoreCase) <= 0) continue;    // el cambio de este PC es más reciente: se enviará
                            st.Pend.Remove(d.Uid);
                            App.Log(string.Format("firebase: {0} cambiado en 2 PC, gana el de {1} ({2})", d.Id, d.By, d.Updated));
                        }
                        // las contraseñas no vienen de Firebase: se queda la de este PC
                        Dictionary<string, object> FV = new Dictionary<string, object>();
                        foreach (Field f in t.Fields) FV[f.Key] = f.Local ? (c != null ? Store.Get(c.F, f.Key) : null) : Store.FromStoreValue(f, Json.Get(o, f.Key));
                        Store.UpdateReqDeadline(t, FV);    // el plazo se calcula aquí (otro PC puede mandar solo las fechas)
                        int row = Row(t, s, c, d.Id, st);
                        if (row == 0) throw new GestemException(Tx.L("errFull"));
                        if (c == null) { c = new ClientRec(); c.Uid = d.Uid; c.New = false; s.Clients.Add(c); }
                        // Sync = true: el Excel lo actualiza el PC donde se hizo el cambio
                        c.Row = row; c.Id = d.Id; c.F = FV; c.Updated = d.Updated; c.By = d.By; c.Sync = true;
                        c.Del = d.Deleted ? new DelMark { By = d.By, At = d.Updated } : null;
                        dirty = true;
                        App.Log(string.Format("firebase: recibido {0} de {1}{2}", d.Id, d.Pc, d.Deleted ? " (eliminado)" : ""));
                    }
                    if (dirty)
                    {
                        FbState.Applying = true; Store.Write(t, s);
                        try { Docs.RepairFolders(t, s); } catch (Exception e) { App.Log("documentos: " + e.Message); }    // nº que cambiaron en otro PC
                    }
                }
                finally { FbState.Applying = false; Store.ExitLock(lk); }
            }
            st.Save();
        }

        // lo que hay que enviar (Step-FbSync, antes de lanzar la vuelta)
        public static List<FbPush> BuildPush(FbState st, out bool drop)
        {
            List<FbPush> push = new List<FbPush>(); drop = false;
            List<string> order = new List<string>(); Dictionary<string, List<string>> groups = new Dictionary<string, List<string>>(StringComparer.CurrentCultureIgnoreCase);
            foreach (DictionaryEntry kv in st.Pend) { string tt = ((PendItem)kv.Value).T; if (!groups.ContainsKey(tt)) { groups[tt] = new List<string>(); order.Add(tt); } groups[tt].Add((string)kv.Key); }
            foreach (string tid in order)
            {
                Tramite t = null; foreach (Tramite x in Tramite.All) if (string.Equals(x.Id, tid, StringComparison.CurrentCultureIgnoreCase)) { t = x; break; }
                StoreData s = t != null ? Store.Read(t) : null;
                if (s == null) continue;    // sin datos de este trámite ahora mismo: se queda pendiente
                Hashtable byUid = PsH.New(); foreach (ClientRec c in s.Clients) if (!string.IsNullOrEmpty(c.Uid)) byUid[c.Uid] = c;
                foreach (string uid in groups[tid])
                {
                    ClientRec c = byUid[uid] as ClientRec;
                    HandDelRec e = null; if (c == null) foreach (HandDelRec x in s.HandDel) if (x != null && string.Equals(x.Uid, uid, StringComparison.CurrentCultureIgnoreCase)) { e = x; break; }
                    PendItem pi = (PendItem)st.Pend[uid];
                    if (e != null)
                    {
                        // cliente de una pestaña hecha a mano eliminado (o vuelto a poner): solo el nombre
                        push.Add(new FbPush { Uid = uid, T = t.Id, Id = "", Datos = Json.WriteCompact(D("name", e.Name)), Updated = Ps.Str(e.At), By = Ps.Str(e.By), H = pi.H, N = false, Deleted = e.On });
                        continue;
                    }
                    if (c == null) { st.Pend.Remove(uid); drop = true; continue; }    // ya no existe (p. ej. pasó a otra fila)
                    push.Add(new FbPush { Uid = uid, T = t.Id, Id = Ps.Str(c.Id), Datos = Json.WriteCompact(StoreFields(t, c)), Updated = Ps.Str(c.Updated), By = Ps.Str(c.By), H = pi.H, N = pi.N, Deleted = c.Del != null });
                }
            }
            return push;
        }

        // lo que vuelve de la vuelta (Step-FbSync, al principio)
        public static void ApplyResult(FbResult r, bool join)
        {
            if (r.Auth != null) Auth = r.Auth;
            if (Ps.Str(r.Err) != LastErr) { LastErr = Ps.Str(r.Err); App.Log(r.Err.Length > 0 ? "firebase: " + r.Err : "firebase: conectado"); }
            if (join) { if (r.Err.Length == 0) Join(r.Pulled, r.Mod); return; }
            if (r.Pulled.Count > 0) Receive(r.Pulled, false);
            FbState st = FbState.Load();
            foreach (FbPush p in r.Pushed) if (st.Pend.ContainsKey(p.Uid) && string.Equals(((PendItem)st.Pend[p.Uid]).H, p.H, StringComparison.CurrentCultureIgnoreCase)) st.Pend.Remove(p.Uid);
            if (r.Mod.Length > 0) st.LastMod = r.Mod;
            if (st.Joined == "push")
            {
                // lo suyo no se pudo subir (p. ej. otro PC subió antes el mismo nº): se vuelve a unir, recibiendo primero
                bool anyNew = false; foreach (DictionaryEntry kv in st.Pend) if (((PendItem)kv.Value).N) anyNew = true;
                if (r.Err.Length > 0) { st.Joined = ""; App.Log("firebase: unión no terminada, se repite"); }
                else if (!anyNew) { st.Joined = "yes"; App.Log("firebase: este PC ya está unido"); }
            }
            st.Save();
            if (r.Pushed.Count > 0) App.Log(string.Format("firebase: enviados {0} cliente(s)", r.Pushed.Count));
        }

        // Set-FbSyncSwitch (casilla de admin). Encender solo con --real (o en la prueba de red, que solo puede usar _tests/)
        public static void SetSwitch(bool on)
        {
            if (on && !App.Real && !TestRun) throw new InvalidOperationException("modo prueba: la sincronización con Firebase no se puede encender");
            Config.Cur.FbSync = on; Config.Save();
            string f = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(App.CfgPath), "firebase_apagado.txt");
            App.Guard(f, "firebase");
            if (on) { try { System.IO.File.Delete(f); } catch { } }
            else System.IO.File.WriteAllText(f, "GESTEM no sincroniza con Firebase en este PC (Ajustes de admin).\r\n", new UTF8Encoding(true));    // Set-Content -Encoding UTF8
        }

        // Step-FbSync (cada 20 s): recoge el resultado de la vuelta anterior y lanza la siguiente en segundo plano.
        // Nunca en modo prueba; TestRun lo pone solo GESTEM.exe --fbtest, y aun así CheckName no deja salir de gestorias/elmahi/_tests/
        public static bool TestRun;
        static Task<FbResult> task; static bool taskJoin;
        public static bool Running { get { return task != null && !task.IsCompleted; } }
        public static void Step()
        {
            if (!Config.Cur.FbSync || !(App.Real || TestRun)) return;
            if (task != null)
            {
                if (!task.IsCompleted) return;
                Task<FbResult> j = task; task = null;
                if (j.Status == TaskStatus.RanToCompletion && j.Result != null) ApplyResult(j.Result, taskJoin);
            }
            string k = Updater.GetKey(); if (k.Length == 0) return;
            FbState st = FbState.Load();
            // primera vez de este PC (o unión sin terminar): solo se recibe todo; no se envía nada hasta compararlo
            bool join = st.Joined != "push" && st.Joined != "yes";
            List<FbPush> push = new List<FbPush>(); bool drop = false;
            if (!join) push = BuildPush(st, out drop);
            if (drop) st.Save();
            FbAuth auth = Auth; string lastMod = join ? "" : st.LastMod, pc = Environment.MachineName;
            taskJoin = join;
            task = Task.Run(delegate { return SyncOnce(k, auth, lastMod, push, pc); });
        }
    }

    public partial class MainWin
    {
        DispatcherTimer liveTimer; int liveTick;
        public DispatcherTimer LiveTimer { get { return liveTimer; } }    // pruebas

        // al abrir (después de cargar): cada 20 s se mira si otro PC guardó y se sincroniza con Firebase
        void StartLiveTimer()
        {
            Fb.Notify = delegate(string m) { ShowToast(m, "", null, 15); };
            if (liveTimer != null) return;
            liveTimer = new DispatcherTimer(); liveTimer.Interval = TimeSpan.FromSeconds(20);
            liveTimer.Tick += delegate { InvokeLiveRefresh(); };
            liveTimer.Start();
        }

        // Invoke-LiveRefresh
        public void InvokeLiveRefresh()
        {
            if (Busy || App.User.Length == 0) return;
            Busy = true;
            try
            {
                liveTick++; bool own = false;
                try { Hand.SyncTabs(false); } catch (Exception e) { App.Log("pestañas propias: " + e.Message); }    // ¿se guardó el Excel? se vuelven a leer sus pestañas
                try { Fb.Step(); } catch (Exception e) { App.Log("firebase: " + e.Message); }
                if (liveTick % 3 == 0) { try { SyncProfiles(false); } catch (Exception e) { App.Log("usuarios (Firebase): " + e.Message); } }    // contraseñas, cada minuto    // lo recibido cambia los datos: la lista se recarga abajo
                // cada 3 min, lo pendiente del Excel. shortcut: sin Test-ExcelFree; hoy no hay nada pendiente (Data.ExcelWrites = false), portarlo si vuelve la escritura
                if (liveTick % 9 == 0 && Data.PendingSyncCount() > 0) { if (Data.SyncPendingToExcel() > 0) { own = true; G.CacheStamp = -1; } }
                if (Store.FileStamp() != G.CacheStamp)
                {
                    ClientRow sel = U<DataGrid>("ClientsGrid").SelectedItem as ClientRow; int selRow = sel != null ? sel.Row : 0;
                    if (Data.LoadData(true, true))
                    {
                        DataGrid g = U<DataGrid>("ClientsGrid");
                        if (selRow > 0) foreach (object it in g.Items) { ClientRow r = it as ClientRow; if (r != null && r.Row == selRow) { g.SelectedItem = it; g.ScrollIntoView(it); break; } }
                        if (U("PanelStats").Visibility == Visibility.Visible) UpdateStats();
                        if (!own) SetStatus(L("liveUpdated"));
                    }
                }
            }
            catch (Exception e) { App.Log("DEBUG actualización automática: " + e.Message); }
            finally { Busy = false; }
        }
    }
}
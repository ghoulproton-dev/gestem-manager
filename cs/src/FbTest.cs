// GESTEM.exe --fbtest <carpeta> <gestoría de prueba>: una vuelta de verdad por Firestore, SOLO en gestorias/elmahi/_tests/<...>.
// Con una copia de certificado.json (tiene un campo Local = contraseña): otro PC sube 2 clientes, este PC se une, envía lo suyo,
// otro PC lo recibe todo y luego nada más; al final se borra TODO lo de la carpeta de prueba y se comprueba que queda vacía.
// También la vuelta de cada 20 s de la ventana (Fb.Step) con el interruptor de admin apagado y encendido.
// Resultado: <carpeta>\cs_fb.txt y la lista de rutas tocadas (cs_fb_paths.txt).
// No escribe nada fuera de <carpeta>: la clave de GESTEM se LEE de la config de este PC (no se guarda nada en ella).
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Gestem
{
    public static class FbTest
    {
        public static int Run(string dir, string gest)
        {
            StringBuilder o = new StringBuilder(); List<string> touched = new List<string>();
            if (!gest.StartsWith("elmahi/_tests/", StringComparison.Ordinal)) throw new InvalidOperationException("las pruebas solo pueden usar gestorias/elmahi/_tests: " + gest);
            string realCfg = App.CfgPath;
            App.Quiet = true; App.AddSafeRoot(dir); App.Real = false;
            App.DataDir = dir; App.BackupDir = Path.Combine(dir, "copias"); App.LogPath = Path.Combine(dir, "cs_fb_log.txt"); App.CfgPath = Path.Combine(dir, "config_fb.json"); FbState.Path = Path.Combine(dir, "cs_fb_state.json");
            Config.Cur = new Config();
            object real = Json.Parse(File.ReadAllText(realCfg, Encoding.UTF8));
            Config.Cur.UpdateKey = Json.Str(real, "UpdateKey") ?? ""; Config.Cur.FbSync = true;
            string sd = Path.Combine(dir, "cs_fb", "GESTEM datos"); if (Directory.Exists(Path.Combine(dir, "cs_fb"))) Directory.Delete(Path.Combine(dir, "cs_fb"), true); Directory.CreateDirectory(sd);
            // la copia de prueba (test.bat), y con nombres inventados: a Firestore no sube ningún cliente de verdad, ni en _tests
            File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test", "cs", "GESTEM datos", "certificado.json"), Path.Combine(sd, "certificado.json"));
            Config.Cur.ExcelPath = Path.Combine(dir, "cs_fb", "x.xlsx");
            Tramite t = Tramite.ById("certificado");
            {
                StoreData s0 = Store.Read(t); List<ClientRec> keep = new List<ClientRec>();
                foreach (ClientRec c in s0.Clients) if (c.Row < Hand.RowBase && keep.Count < 3) keep.Add(c);
                for (int i = 0; i < keep.Count; i++) { Dictionary<string, object> F = new Dictionary<string, object>(); F["name"] = "PRUEBA FIREBASE " + (i + 1); F["xH"] = "CLAVE-" + (i + 1); keep[i].F = F; }
                s0.Clients = keep; s0.HandDel.Clear(); s0.Adopted.Clear();
                FbState.Applying = true; try { Store.Write(t, s0); } finally { FbState.Applying = false; }
            }
            if (File.Exists(FbState.Path)) File.Delete(FbState.Path);
            Fb.Gest = gest; Fb.Auth = null; Fb.LastErr = "";
            Fb.Trace = delegate(string m, string p) { string k = m + " " + p.Replace(Fb.Root + "/", ""); if (!touched.Contains(k)) touched.Add(k); };
            string key = Updater.GetKey(); if (key.Length == 0) throw new Exception("falta la clave de GESTEM en la config de este PC");
            FbAuth a = Fb.SignIn(key);
            try
            {
                List<string> before = Fb.ListAll(a);
                if (before.Count > 0) { o.Append("ABORT la carpeta de prueba no está vacía: " + before.Count + " documento(s)\r\n"); File.WriteAllText(Path.Combine(dir, "cs_fb.txt"), o.ToString(), new UTF8Encoding(false)); return 3; }
                o.Append("F0 vacía=True\r\n");
                // otro PC sube 2 clientes de certificado: uno es el mismo que el primero de este PC (mismo nº y nombre, otro estado), otro es nuevo
                StoreData s = Store.Read(t); ClientRec c0 = null; foreach (ClientRec c in s.Clients) if (c.Row < Hand.RowBase) { c0 = c; break; }
                Dictionary<string, object> f0 = Fb.StoreFields(t, c0); f0["xH"] = "NO-DEBE-LLEGAR";    // aunque otro PC lo mande, no se aplica (Local)
                f0["status"] = "DESDE OTRO PC";
                Dictionary<string, object> f1 = new Dictionary<string, object>(); f1["name"] = "ZZ PRUEBA FIREBASE"; f1["xH"] = "NO-DEBE-LLEGAR";
                string[] uids = { "otro-1", "otro-2" }; string[] idsx = { c0.Id, "CER-2026-090" }; object[] datos = { f0, f1 };
                List<object> w = new List<object>(); List<string> names = new List<string>();
                for (int i = 0; i < 2; i++)
                {
                    string name = Fb.Base + "/clientes/" + uids[i]; names.Add(name);
                    Dictionary<string, object> fields = new Dictionary<string, object>();
                    fields["t"] = SV(t.Id); fields["id"] = SV(idsx[i]); fields["datos"] = SV(Json.WriteCompact(datos[i])); fields["updated"] = SV("2099-01-01T00:00:0" + i); fields["by"] = SV("OTRA"); fields["pc"] = SV("OTROPC");
                    Dictionary<string, object> del = new Dictionary<string, object>(); del["booleanValue"] = false; fields["deleted"] = del;
                    Dictionary<string, object> upd = new Dictionary<string, object>(); upd["name"] = name; upd["fields"] = fields;
                    Dictionary<string, object> tr = new Dictionary<string, object>(); tr["fieldPath"] = "mod"; tr["setToServerValue"] = "REQUEST_TIME";
                    Dictionary<string, object> wr = new Dictionary<string, object>(); wr["update"] = upd; wr["updateTransforms"] = new List<object> { tr };
                    w.Add(wr);
                }
                Fb.Commit(a, w, names);
                { int seen = 0; foreach (string n in Fb.ListUnderTests(a)) if (names.Contains(n)) seen++; o.Append("F0b limpieza ve=" + seen + "/2\r\n"); }
                // este PC: unión (recibe todo), envío, y otra vuelta sin nada
                FbResult r1 = Fb.SyncOnce(key, a, "", new List<FbPush>(), Environment.MachineName);
                o.Append("F1 err=" + r1.Err + " recibidos=" + r1.Pulled.Count + " mod=" + (r1.Mod.Length > 0) + "\r\n");
                Fb.ApplyResult(r1, true);
                FbState st = FbState.Load(); bool drop; List<FbPush> push = Fb.BuildPush(st, out drop);
                int bad = 0; foreach (FbPush p in push) foreach (Field f in t.Fields) if (f.Local && p.Datos.Contains("\"" + f.Key + "\":")) bad++;
                o.Append("F2 joined=" + st.Joined + " porEnviar=" + push.Count + " conContraseña=" + bad + "\r\n");
                FbResult r2 = Fb.SyncOnce(key, a, st.LastMod, push, Environment.MachineName);
                o.Append("F3 err=" + r2.Err + " enviados=" + r2.Pushed.Count + " recibidos=" + r2.Pulled.Count + "\r\n");
                Fb.ApplyResult(r2, false); st = FbState.Load();
                o.Append("F4 joined=" + st.Joined + " pend=" + st.Pend.Count + "\r\n");
                foreach (ClientRec c in Store.Read(t).Clients) if (c.Uid == "otro-1" || c.Uid == "otro-2") o.Append("F5 " + c.Uid + " " + c.Id + " " + Ps.Str(Store.Get(c.F, "name")) + "|" + Ps.Str(Store.Get(c.F, "status")) + "|xH=" + Ps.Str(Store.Get(c.F, "xH")) + "\r\n");
                // otro PC lo recibe todo; nada de lo que hay en Firestore lleva contraseñas; desde la última vez, nada más
                FbResult r3 = Fb.SyncOnce(key, a, "", new List<FbPush>(), "OTROPC");
                int bad2 = 0, fromMe = 0; foreach (FbDoc d in r3.Pulled) { foreach (Field f in t.Fields) if (f.Local && d.Pc == Environment.MachineName && d.Datos.Contains("\"" + f.Key + "\":")) bad2++; if (d.Pc == Environment.MachineName) fromMe++; }
                o.Append("F6 err=" + r3.Err + " total=" + r3.Pulled.Count + " deEstePC=" + fromMe + " conContraseña=" + bad2 + "\r\n");
                FbResult r4 = Fb.SyncOnce(key, a, r3.Mod, new List<FbPush>(), "OTROPC");
                o.Append("F7 err=" + r4.Err + " nuevos=" + r4.Pulled.Count + "\r\n");
                // la vuelta de cada 20 s de la ventana (Fb.Step) con el interruptor de admin: apagado no hace nada; encendido envía un cambio de este PC
                Fb.TestRun = true; Fb.Auth = a;
                Fb.SetSwitch(false); Fb.Step();
                o.Append("F10 apagado enMarcha=" + Fb.Running + " archivo=" + File.Exists(Path.Combine(dir, "firebase_apagado.txt")) + "\r\n");
                Fb.SetSwitch(true);
                { StoreData s2 = Store.Read(t); foreach (ClientRec c in s2.Clients) if (c.Row < Hand.RowBase && c.Uid != "otro-1" && c.Uid != "otro-2") { c.F["status"] = "CAMBIO 20 S"; break; } Store.Write(t, s2); }
                for (int i = 0; i < 3; i++) { Fb.Step(); for (int wt = 0; wt < 600 && Fb.Running; wt++) System.Threading.Thread.Sleep(100); }    // envía, recoge, recoge la vuelta vacía
                FbResult r5 = Fb.SyncOnce(key, a, r4.Mod.Length > 0 ? r4.Mod : r3.Mod, new List<FbPush>(), "OTROPC");    // F7 no recibió nada: su Mod viene vacío
                string got = ""; foreach (FbDoc d in r5.Pulled) if (d.Datos.Contains("CAMBIO 20 S")) got = "CAMBIO 20 S";
                o.Append("F11 pend=" + FbState.Load().Pend.Count + " otroRecibe=" + r5.Pulled.Count + " estado=" + got + " err=" + Fb.LastErr + "\r\n");
                // usuarios: en Firestore solo hash, sal y estado (ni la contraseña ni el icono); otro PC la recibe y la restablece; este lo recibe
                {
                    string keepData = App.DataDir, otro = Path.Combine(dir, "cs_fb_otro_pc"); if (Directory.Exists(otro)) Directory.Delete(otro, true); Directory.CreateDirectory(otro);
                    foreach (string f in new string[] { "usuarios.json" }) if (File.Exists(Path.Combine(dir, f))) File.Delete(Path.Combine(dir, f));
                    Profiles.Reload(); Profiles.SetIcon("PRUEBA USUARIO", "#15803D", ""); Profiles.SetPassword("PRUEBA USUARIO", "clave-prueba-1");
                    FbAuth ua = a; List<Profile> u1 = Profiles.SyncOnce(key, ref ua, Profiles.Snapshot());
                    object raw = Fb.J("POST", "https://firestore.googleapis.com/v1/" + Fb.Base + ":runQuery", Fb.D("structuredQuery", Fb.D("from", new List<object> { Fb.D("collectionId", "usuarios") })), a.Id, null);
                    string rt = Json.WriteCompact(raw);
                    o.Append("F12 enFirebase=" + u1.Count + " hash=" + rt.Contains(Profiles.Get("PRUEBA USUARIO").Hash) + " contraseña=" + rt.Contains("clave-prueba-1") + " icono=" + rt.Contains("15803D") + "\r\n");
                    App.DataDir = otro; Profiles.Reload();
                    bool recv = Profiles.Merge(Profiles.SyncOnce(key, ref ua, Profiles.Snapshot()));
                    bool ok = Profiles.Check("PRUEBA USUARIO", "clave-prueba-1");
                    Profiles.Reset("PRUEBA USUARIO", "admin"); Profiles.SyncOnce(key, ref ua, Profiles.Snapshot());
                    o.Append("F13 otroPC recibe=" + recv + " entra=" + ok + " icono=" + (Profiles.Get("PRUEBA USUARIO").Color == "" ? "no" : "sí") + "\r\n");
                    App.DataDir = keepData; Profiles.Reload();
                    bool back = Profiles.Merge(Profiles.SyncOnce(key, ref ua, Profiles.Snapshot()));
                    o.Append("F14 esteRecibeRestablecida=" + back + " tieneContraseña=" + Profiles.HasPassword("PRUEBA USUARIO") + " icono=" + Profiles.Get("PRUEBA USUARIO").Color + "\r\n");
                }
                // gestoría de verdad: la puerta no deja ni mirarla sin --real
                string keep = Fb.Gest; Fb.Gest = "elmahi"; try { Fb.CheckName(Fb.Base + "/clientes/x"); o.Append("F8 puerta=MAL\r\n"); } catch (InvalidOperationException) { o.Append("F8 puerta=OK\r\n"); } finally { Fb.Gest = keep; }
                try { Fb.CheckName(Fb.Root + "/gestorias/elmahi/clientes/x"); o.Append("F9 fuera=MAL\r\n"); } catch (InvalidOperationException) { o.Append("F9 fuera=OK\r\n"); }
            }
            finally
            {
                // limpieza: todo lo que haya debajo de la carpeta de prueba (también si algo falló a medias)
                int n = 0; foreach (string d in Fb.ListAll(a)) { Fb.DeleteDoc(a, d); n++; }
                o.Append("FZ borrados=" + n + " quedan=" + Fb.ListAll(a).Count + "\r\n");
                Fb.Trace = null; Fb.TestRun = false;
                touched.Sort(StringComparer.Ordinal);
                File.WriteAllText(Path.Combine(dir, "cs_fb.txt"), o.ToString(), new UTF8Encoding(false));
                File.WriteAllLines(Path.Combine(dir, "cs_fb_paths.txt"), touched.ToArray(), new UTF8Encoding(false));
            }
            return 0;
        }
        // cada carpeta de gestorias/elmahi/_tests: todos sus documentos (y ella misma si existe) se borran; resultado en cs_fbclean.txt
        public static int Clean(string dir)
        {
            StringBuilder o = new StringBuilder(); List<string> touched = new List<string>();
            string realCfg = App.CfgPath; App.Quiet = true; App.Real = false; App.AddSafeRoot(dir);
            App.DataDir = dir; App.BackupDir = Path.Combine(dir, "copias"); App.LogPath = Path.Combine(dir, "cs_fbclean_log.txt"); App.CfgPath = Path.Combine(dir, "config_fb.json");
            Config.Cur = new Config(); Config.Cur.UpdateKey = Json.Str(Json.Parse(File.ReadAllText(realCfg, Encoding.UTF8)), "UpdateKey") ?? "";
            string key = Updater.GetKey(); if (key.Length == 0) throw new Exception("falta la clave de GESTEM en la config de este PC");
            Fb.Trace = delegate(string m, string p) { string k = m + " " + p.Replace(Fb.Root + "/", ""); if (!touched.Contains(k)) touched.Add(k); };
            string keep = Fb.Gest;
            try
            {
                Fb.Gest = "elmahi/_tests/x"; FbAuth a = Fb.SignIn(key);
                List<string> docs = Fb.ListUnderTests(a);
                o.Append("encontrados=" + docs.Count + "\r\n");
                foreach (string d in docs)
                {
                    string rel = d.Substring((Fb.Root + "/gestorias/elmahi/_tests/").Length);
                    Fb.Gest = "elmahi/_tests/" + rel.Split('/')[0];    // Fb.CheckName: solo dentro de esa carpeta de prueba
                    Fb.DeleteDoc(a, d);
                    o.Append("borrado " + d.Replace(Fb.Root + "/", "") + "\r\n");
                }
                Fb.Gest = "elmahi/_tests/x"; o.Append("quedan=" + Fb.ListUnderTests(a).Count + "\r\n");
            }            finally
            {
                Fb.Gest = keep; Fb.Trace = null; touched.Sort(StringComparer.Ordinal);
                File.WriteAllText(Path.Combine(dir, "cs_fbclean.txt"), o.ToString(), new UTF8Encoding(false));
                File.WriteAllLines(Path.Combine(dir, "cs_fbclean_paths.txt"), touched.ToArray(), new UTF8Encoding(false));
            }
            return 0;
        }
        static Dictionary<string, object> SV(string v) { Dictionary<string, object> d = new Dictionary<string, object>(); d["stringValue"] = v; return d; }
    }
}

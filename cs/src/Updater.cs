// Etapa 5b: actualizaciones y control remoto. Copiado de Get-GestemKey, Set-GestemKey, Get-RepoFile, Test-GestemKey, Ask-GestemKey,
// Start-UpdateCheckAsync, Check-Update y Check-Remote de GESTEM.ps1. Solo LEE de GitHub (version.txt, control.json): nunca escribe en
// el repositorio, ni toca GESTEM.ps1 ni version.txt.
// Etapa 7: GESTEM.exe mira app/exe_version.txt (version.txt es de GESTEM.ps1) y se instala solo (ExeSwitch.Update).
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Gestem
{
    public class RepoException : Exception { public int Code; public RepoException(int code, string msg) : base(msg) { Code = code; } }

    public static class Updater
    {
        public const string UpdateBase = "https://raw.githubusercontent.com/ghoulproton-dev/gestem-manager/main";
        public const string RepoApi = "https://api.github.com/repos/ghoulproton-dev/gestem-manager";
        public static Func<string, string> FetchHook;    // pruebas: en vez de GitHub (lanza RepoException para un 404…)
        public static Func<string, byte[]> FetchBytesHook;    // pruebas: archivos binarios (app/GESTEM.exe)
        public const string ExeVersionFile = "app/exe_version.txt";

        // clave de GESTEM: abre el GitHub privado (cifrada para este usuario de Windows)
        public static string GetKey()
        {
            string e = Ps.Str(Config.Cur.UpdateKey); if (e.Length == 0) return "";
            try { return Encoding.UTF8.GetString(System.Security.Cryptography.ProtectedData.Unprotect(Convert.FromBase64String(e), null, System.Security.Cryptography.DataProtectionScope.CurrentUser)); } catch { return ""; }
        }
        public static void SetKey(string k)
        {
            Config.Cur.UpdateKey = string.IsNullOrEmpty(k) ? "" : Convert.ToBase64String(System.Security.Cryptography.ProtectedData.Protect(Encoding.UTF8.GetBytes(k), null, System.Security.Cryptography.DataProtectionScope.CurrentUser));
            Config.Save();
        }

        static HttpClient Client(string key, int timeout)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            HttpClient hc = new HttpClient(); hc.Timeout = TimeSpan.FromSeconds(timeout);
            hc.DefaultRequestHeaders.Add("User-Agent", "GESTEM");
            if (key.Length > 0) { hc.DefaultRequestHeaders.Add("Authorization", "Bearer " + key); hc.DefaultRequestHeaders.Add("Accept", "application/vnd.github.raw"); }
            return hc;
        }
        static string Uri(string name, string key) { return key.Length > 0 ? RepoApi + "/contents/" + name + "?ref=main" : UpdateBase + "/" + name + "?t=" + DateTime.UtcNow.Ticks; }

        // con clave: API de GitHub (repositorio privado); sin clave: el archivo público. Solo lectura.
        public static string RepoFile(string name, int timeout = 5)
        {
            if (FetchHook != null) return FetchHook(name);
            return Encoding.UTF8.GetString(RepoBytes(name, timeout));
        }
        public static byte[] RepoBytes(string name, int timeout)
        {
            if (FetchBytesHook != null) return FetchBytesHook(name);
            string k = GetKey();
            using (HttpClient hc = Client(k, timeout))
            {
                HttpResponseMessage r;
                try { r = hc.GetAsync(Uri(name, k)).Result; }
                catch (AggregateException e) { throw new RepoException(0, e.InnerException != null ? e.InnerException.Message : e.Message); }
                using (r)
                {
                    int code = (int)r.StatusCode;
                    if (!r.IsSuccessStatusCode)
                    {
                        if (k.Length > 0 && code == 401) { SetKey(""); App.Log("clave de GESTEM rechazada por GitHub: se pedirá otra"); }
                        throw new RepoException(code, "(" + code + ") " + r.ReasonPhrase);
                    }
                    return r.Content.ReadAsByteArrayAsync().Result;
                }
            }
        }
        public static Task<string> VersionAsync()
        {
            if (FetchHook != null) return Task.FromResult(FetchHook(ExeVersionFile));
            string k = GetKey(); HttpClient hc = Client(k, 20);
            return hc.GetStringAsync(Uri(ExeVersionFile, k)).ContinueWith(t => { hc.Dispose(); return t.Result; });
        }

        public static string TestKey(string k)
        {
            try
            {
                using (HttpClient hc = Client("", 10))
                {
                    hc.DefaultRequestHeaders.Add("Authorization", "Bearer " + k);
                    using (HttpResponseMessage r = hc.GetAsync(RepoApi).Result)
                    {
                        int code = (int)r.StatusCode;
                        if (r.IsSuccessStatusCode) return "ok";
                        return code == 401 || code == 403 || code == 404 ? "bad" : "offline";
                    }
                }
            }
            catch { return "offline"; }
        }

        // «1.9.4» > «1.9.3»; lo que no es un número de versión no cuenta
        public static bool IsNewer(string remote, string current)
        {
            remote = Ps.Str(remote).Trim();
            return Regex.IsMatch(remote, @"^\d+(\.\d+){1,3}$") && new Version(remote) > new Version(current);
        }
    }

    public partial class MainWin
    {
        bool updChecking;
        bool leaving;    // se cierra para volver a PowerShell (o para abrir el GESTEM.exe nuevo): no se carga nada más
        public Action<string> InfoHook;    // pruebas: qué mensaje de información se mostró

        void WireUpdater()
        {
            U<Button>("UpdateBtn").Click += delegate { Win.Cursor = Cursors.Wait; CheckUpdate(true); Win.Cursor = null; };
        }

        // al abrir: clave (una vez por PC), versión nueva en segundo plano, control remoto
        void StartupOnline()
        {
            if (Updater.GetKey().Length == 0) { try { AskGestemKey(); } catch (Exception e) { App.Log("clave de GESTEM: " + e.Message); } }
            StartUpdateCheckAsync();
            CheckRemote();
        }

        void AskGestemKey()
        {
            for (int i = 0; i < 3; i++)
            {
                string k = Ps.Str(Microsoft.VisualBasic.Interaction.InputBox(L("keyAsk"), App.Name, "")).Trim();
                if (k.Length == 0) return;
                string r = Updater.TestKey(k);
                if (r == "ok") { Updater.SetKey(k); App.Log("clave de GESTEM guardada"); Toast(L("keyOk"), false, 7); return; }
                if (r == "offline") { ShowError(L("keyOffline")); return; }
                ShowError(L("keyBad"));
            }
        }

        // mira si hay versión nueva sin congelar la ventana
        async void StartUpdateCheckAsync()
        {
            string remote;
            try { remote = Ps.Str(await Updater.VersionAsync()).Trim(); }
            catch (Exception e) { Exception x = e is AggregateException && e.InnerException != null ? e.InnerException : e; App.Log("actualización: no se pudo mirar (" + x.Message + ")"); return; }
            if (!Updater.IsNewer(remote, App.Version)) return;
            App.Log("actualización: hay versión " + remote + " (esta: " + App.Version + ")");
            if (!FormHasWork()) CheckUpdate(false);
            else { App.Log("actualización " + remote + " lista: espera (ficha a medias)"); ShowToast(string.Format(L("updReady"), remote), L("updNow"), delegate { CheckUpdate(true); }, 30); }
        }

        public void CheckUpdate(bool manual)
        {
            if (updChecking) { App.Log("actualización: ya se está mirando, se omite"); return; }
            updChecking = true;
            try
            {
                string remote = Ps.Str(Updater.RepoFile(Updater.ExeVersionFile)).Trim();
                if (!Regex.IsMatch(remote, @"^\d+(\.\d+){1,3}$")) throw new Exception("version.txt no válido: " + remote);
                if (!Updater.IsNewer(remote, App.Version)) { if (manual) Toast(string.Format(L("upToDate"), App.Version), false, 7); return; }
                if (!ExeSwitch.CanUpdate)
                {
                    App.Log("actualización: hay GESTEM.exe " + remote + "; esta copia no se actualiza sola (modo prueba o sin GESTEM.ps1 al lado)");
                    if (manual) ShowError(L("updateFail") + Environment.NewLine + "GESTEM.exe " + App.Version + " → " + remote + ": esta copia no se actualiza sola.");
                    return;
                }
                SetStatus(string.Format(L("updating"), remote));
                ExeSwitch.Update(remote);
                // con una ficha a medias no se cierra nada: el nuevo se abre la próxima vez
                if (FormHasWork()) { Toast(string.Format(L("updated"), remote), false, 10); return; }
                leaving = true; ExeSwitch.Restart(); Win.Close();
            }
            catch (Exception e)
            {
                App.Log("actualización: " + e.Message);
                if (manual) ShowError(L("updateFail") + Environment.NewLine + e.Message);
            }
            finally { updChecking = false; }
        }

        // control.json de GitHub: mensaje (una vez) y pausa remota; sin conexión se usa la última copia
        public void CheckRemote()
        {
            string cache = Path.Combine(App.DataDir, "control.cache");
            object j = null;
            try
            {
                string txt = Updater.RepoFile("control.json");
                j = Json.Parse(txt);
                App.Guard(cache, "control remoto");
                File.WriteAllText(cache, txt + "\r\n", new UTF8Encoding(true));    // Set-Content -Encoding UTF8
            }
            catch (Exception e)
            {
                RepoException re = e as RepoException;
                if (re != null && re.Code == 404) { if (File.Exists(cache)) { App.Guard(cache, "control remoto"); File.Delete(cache); } }
                else if (File.Exists(cache)) { try { j = Json.Parse(File.ReadAllText(cache, Encoding.UTF8)); } catch { } }
            }
            Remote = j;
            if (j == null) return;
            if (App.Real && ExeSwitch.WantsPs(j)) { leaving = true; ExeSwitch.BackToPs(); Win.Close(); return; }    // vuelta remota a PowerShell
            string msg = Ps.Str(Json.Str(j, "mensaje")).Trim();
            if (msg.Length > 0)
            {
                string seenPath = Path.Combine(App.DataDir, "control.visto");
                string seen = File.Exists(seenPath) ? File.ReadAllText(seenPath, Encoding.UTF8).Trim() : "";
                if (seen != msg) { ShowInfo(msg); App.Guard(seenPath, "control remoto"); File.WriteAllText(seenPath, msg + "\r\n", new UTF8Encoding(true)); }
            }
            ApplyPause();
        }

        void ShowInfo(string msg) { if (InfoHook != null) InfoHook(msg); Info(msg); }
    }
}

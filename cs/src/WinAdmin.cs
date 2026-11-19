// Etapa 4d: administrador (oculto, Ctrl+Mayús+Alt+A). Copiado de Test-OwnerPC, Set-AdmTree, Apply-AdminLanguage, Convert-LogText,
// Load-Admin, Save-Admin, Get-PwHash, Get-LockSet, Test-Locked, Apply-Locks, Show-AdminLogin, Update-KeyHint, Invoke-AdminLogin,
// Open-AdminPanel, Fill-AdminLog, Test-ErrorLogLine, Show-AdminTab, Save-AdminLocks, Update-AdmFbState, Change-AdminPw,
// Close-Admin, Apply-Pause y Update-RemoteText de GESTEM.ps1 (y sus eventos).
// shortcut: leer control.json de GitHub (Check-Remote) va con las actualizaciones (Get-RepoFile, etapa 5): hasta entonces no hay
// pausa ni mensaje remoto. Encender/apagar Firebase (Set-FbSyncSwitch) es de la 5b: la casilla se ve pero no se puede cambiar.
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace Gestem
{
    public class AdminData { public string Salt = "", Hash = ""; public Hashtable Locks = PsH.New(); public bool IgnorePause, OwnerPC; }
    public class AdmLogItem { public string Time { get; set; } public string User { get; set; } public string UserBg { get; set; } public string UserFg { get; set; } public string Text { get; set; } public string Fg { get; set; } }

    public partial class MainWin
    {
        public static AdminData Admin = new AdminData();
        public static string AdminPath { get { return Path.Combine(App.DataDir, "admin.dat"); } }
        // PC del creador: se reconoce solo (huella cifrada de nombre del PC + usuario de Windows, no se guarda el nombre)
        static readonly string[] OwnerPrints = { "799CEE89435398F2A3BAA342547E3E364B49C27853D8398CD2B837E49B83271D" };
        public static readonly string[] LockKeys = { "File", "Ext", "Comp", "Team", "Pref", "Help", "All" };
        public string AdmLogFilter = "all";
        public object Remote;    // $script:Remote (control.json); null hasta la etapa 5
        readonly Dictionary<object, string> trText = new Dictionary<object, string>(), trTip = new Dictionary<object, string>();

        static string AdmT(string es) { return Tx.AdmT(es); }

        void WireAdmin()
        {
            U<Button>("AdmLoginBtn").Click += delegate { InvokeAdminLogin(); };
            foreach (string n in new string[] { "AdmPw", "AdmPw2" }) U<PasswordBox>(n).KeyDown += delegate(object s, KeyEventArgs e) { if (e.Key == Key.Return) InvokeAdminLogin(); };
            foreach (string n in new string[] { "AdmPw", "AdmPw2", "AdmNewPw", "AdmNewPw2" }) { U<PasswordBox>(n).KeyUp += delegate { UpdateKeyHint(); }; U<PasswordBox>(n).GotKeyboardFocus += delegate { UpdateKeyHint(); }; }
            U<Button>("AdmLoginCancel").Click += delegate { U("PanelAdminLogin").Visibility = Visibility.Collapsed; };
            U<Button>("AdmClose").Click += delegate { U("PanelAdmin").Visibility = Visibility.Collapsed; };
            U<Button>("AdmLogout").Click += delegate { CloseAdmin(); ApplyPause(); };
            U<ToggleButton>("AdmTabLocks").Checked += delegate { ShowAdminTab("AdmPageLocks"); };
            U<ToggleButton>("AdmTabLog").Checked += delegate { FillAdminLog(); ShowAdminTab("AdmPageLog"); };
            U<ToggleButton>("AdmTabTools").Checked += delegate { UpdateRemoteText(); ShowAdminTab("AdmPageTools"); };
            U<Button>("AdmSaveLocks").Click += delegate { SaveAdminLocks(); };
            U<Button>("AdmLogRefresh").Click += delegate { FillAdminLog(); };
            U<ToggleButton>("AdmLogAll").Checked += delegate { AdmLogFilter = "all"; FillAdminLog(); };
            U<ToggleButton>("AdmLogTeam").Checked += delegate { AdmLogFilter = "team"; FillAdminLog(); };
            U<ToggleButton>("AdmLogAdmin").Checked += delegate { AdmLogFilter = "admin"; FillAdminLog(); };
            U<ToggleButton>("AdmLogErr").Checked += delegate { AdmLogFilter = "errors"; FillAdminLog(); };
            U<Button>("AdmChangePw").Click += delegate { ChangeAdminPw(); };
            U<Button>("AdmOpenData").Click += delegate { System.Diagnostics.Process.Start("explorer.exe", "\"" + App.DataDir + "\""); };
            U<Button>("AdmOpenApp").Click += delegate { System.Diagnostics.Process.Start("explorer.exe", "\"" + App.AppDir + "\""); };
            U<Button>("AdmShowAll").Click += delegate { U("PanelAdmin").Visibility = Visibility.Collapsed; LoadSettingsPanel(); U<ToggleButton>("NavSettings").IsChecked = true; ShowPanel("PanelSettings"); };
            U("AdmUpdate").Visibility = Visibility.Collapsed;    // buscar actualizaciones: etapa 5
            U("AdmFbSync").IsEnabled = false;                     // Set-FbSyncSwitch: etapa 5b
            LoadAdmin();
            App.IsOwnerPC = TestOwnerPC();
            ApplyLocks();
        }

        public static bool TestOwnerPC()
        {
            if (Admin.OwnerPC) return true;
            try
            {
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] b = sha.ComputeHash(Encoding.UTF8.GetBytes(("GESTEM|" + Environment.MachineName + "|" + Environment.UserName).ToUpperInvariant()));
                    return Array.IndexOf(OwnerPrints, BitConverter.ToString(b).Replace("-", "")) >= 0;
                }
            }
            catch { return false; }
        }

        // textos del área de administrador en el idioma elegido (en la XAML están en español)
        void SetAdmTree(object node)
        {
            DependencyObject d = node as DependencyObject;
            if (d == null) return;
            TextBlock tb = d as TextBlock; ContentControl cc = d as ContentControl;
            if (tb != null)
            {
                if (!trText.ContainsKey(tb) && Tx.HasAdm(tb.Text)) trText[tb] = tb.Text;
                if (trText.ContainsKey(tb)) tb.Text = AdmT(trText[tb]);
            }
            else if (cc != null && cc.Content is string)
            {
                if (!trText.ContainsKey(cc) && Tx.HasAdm((string)cc.Content)) trText[cc] = (string)cc.Content;
                if (trText.ContainsKey(cc)) cc.Content = AdmT(trText[cc]);
            }
            FrameworkElement fe = d as FrameworkElement;
            if (fe != null && fe.ToolTip is string)
            {
                if (!trTip.ContainsKey(fe) && Tx.HasAdm((string)fe.ToolTip)) trTip[fe] = (string)fe.ToolTip;
                if (trTip.ContainsKey(fe)) fe.ToolTip = AdmT(trTip[fe]);
            }
            foreach (object c in LogicalTreeHelper.GetChildren(d)) SetAdmTree(c);
        }
        public void ApplyAdminLanguage() { foreach (string p in new string[] { "PanelPaused", "PanelAdminLogin", "PanelAdmin", "PanelHelp", "PanelAbout" }) SetAdmTree(U(p)); }

        // registro de actividad: traduce solo el principio conocido de cada línea
        static string ConvertLogText(string txt)
        {
            if (Ps.Str(Config.Cur.Language) == "es") return txt;
            foreach (string k in Tx.AdmKeysLongestFirst()) if (k.Length >= 8 && txt.StartsWith(k, StringComparison.CurrentCulture)) return AdmT(k) + txt.Substring(k.Length);
            return txt;
        }

        // admin.dat: cifrado con la cuenta de Windows (DPAPI), oculto
        public static void LoadAdmin()
        {
            if (!File.Exists(AdminPath)) return;
            try
            {
                byte[] enc = File.ReadAllBytes(AdminPath);
                object j = Json.Parse(Encoding.UTF8.GetString(ProtectedData.Unprotect(enc, null, DataProtectionScope.CurrentUser)));
                Admin.Salt = Ps.Str(Json.Get(j, "Salt")); Admin.Hash = Ps.Str(Json.Get(j, "Hash"));
                Admin.IgnorePause = Truthy(Json.Get(j, "IgnorePause")); Admin.OwnerPC = Truthy(Json.Get(j, "OwnerPC"));
                Admin.Locks = PsH.New();
                Dictionary<string, object> l = Json.Obj(Json.Get(j, "Locks"));
                if (l != null) foreach (KeyValuePair<string, object> kv in l) Admin.Locks[kv.Key] = Truthy(kv.Value);
            }
            catch (Exception e) { App.Log("admin: no se pudo leer (" + e.Message + ")"); }
        }
        // [bool] de PowerShell sobre un valor de JSON: texto no vacío = verdadero (también "false"), número distinto de 0
        static bool Truthy(object v) { if (v == null) return false; if (v is bool) return (bool)v; if (v is string) return ((string)v).Length > 0; try { return Convert.ToDouble(v) != 0; } catch { return true; } }
        public static void SaveAdmin()
        {
            App.Guard(AdminPath, "admin");
            Dictionary<string, object> o = new Dictionary<string, object>();
            o["Salt"] = Admin.Salt; o["Hash"] = Admin.Hash; o["Locks"] = Admin.Locks; o["IgnorePause"] = Admin.IgnorePause; o["OwnerPC"] = Admin.OwnerPC;
            byte[] enc = ProtectedData.Protect(Encoding.UTF8.GetBytes(Json.Write(o)), null, DataProtectionScope.CurrentUser);
            if (File.Exists(AdminPath)) File.SetAttributes(AdminPath, FileAttributes.Normal);
            File.WriteAllBytes(AdminPath, enc);
            File.SetAttributes(AdminPath, FileAttributes.Hidden);
        }
        public static string PwHash(string pw, string salt)
        {
            using (SHA256 sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(salt + "|GESTEM|" + pw))).Replace("-", "");
        }
        static bool GetLockSet(string k) { return Admin.Locks.ContainsKey(k) && (bool)Admin.Locks[k]; }
        static bool TestLocked(string k) { return !App.IsOwnerPC && GetLockSet(k); }
        public void ApplyLocks()
        {
            foreach (string k in new string[] { "File", "Ext", "Comp", "Team", "Pref", "Help" })
                U("Sec" + k + "Box").Visibility = TestLocked(k) && !App.IsAdmin ? Visibility.Collapsed : Visibility.Visible;
            bool hideAll = TestLocked("All") && !App.IsAdmin;
            U("NavSettings").Visibility = hideAll ? Visibility.Collapsed : Visibility.Visible;
            if (hideAll && U("PanelSettings").Visibility == Visibility.Visible) U<ToggleButton>("NavNew").IsChecked = true;
        }

        public void ShowAdminLogin()
        {
            bool first = Admin.Hash.Length == 0;
            U<TextBlock>("AdmLoginHint").Text = first ? AdmT("Primera vez: crea tu contraseña de administrador (mínimo 6 caracteres) y escríbela dos veces. Guárdala bien: no se puede recuperar.") : AdmT("Escribe tu contraseña de administrador.");
            U("AdmPw2").Visibility = first ? Visibility.Visible : Visibility.Collapsed;
            U<PasswordBox>("AdmPw").Password = ""; U<PasswordBox>("AdmPw2").Password = ""; U<TextBlock>("AdmLoginMsg").Text = "";
            UpdateKeyHint();
            U("PanelAdminLogin").Visibility = Visibility.Visible;
            U("AdmPw").Focus();
        }
        // aviso de Bloq Mayús y del teclado activo (AZERTY/QWERTY cambian la contraseña)
        void UpdateKeyHint()
        {
            string layout = "";
            try { layout = System.Windows.Forms.InputLanguage.CurrentInputLanguage.LayoutName; } catch { }
            if (string.IsNullOrEmpty(layout)) try { layout = InputLanguageManager.Current.CurrentInputLanguage.DisplayName; } catch { }
            string txt = string.IsNullOrEmpty(layout) ? "" : string.Format(AdmT("Teclado: {0}"), layout);
            bool caps = false; try { caps = Keyboard.IsKeyToggled(Key.CapsLock); } catch { }
            TextBlock h = U<TextBlock>("AdmKeyHint");
            if (caps) { h.Text = AdmT("Bloq Mayús está activado.") + (txt.Length > 0 ? "   ·   " + txt : ""); h.Foreground = Br("#B42318"); }
            else { h.Text = txt; h.Foreground = Br("#8A7F81"); }
        }
        public void InvokeAdminLogin()
        {
            string pw = U<PasswordBox>("AdmPw").Password;
            TextBlock msg = U<TextBlock>("AdmLoginMsg");
            if (Admin.Hash.Length == 0)
            {
                if (pw.Length < 6) { msg.Text = AdmT("Mínimo 6 caracteres."); return; }
                if (pw != U<PasswordBox>("AdmPw2").Password) { msg.Text = AdmT("Las dos contraseñas no son iguales."); return; }
                Admin.Salt = Guid.NewGuid().ToString("N");
                Admin.Hash = PwHash(pw, Admin.Salt);
                SaveAdmin();
            }
            else if (PwHash(pw, Admin.Salt) != Admin.Hash)
            {
                msg.Text = AdmT("Contraseña incorrecta.");
                App.Log("admin: intento de acceso fallido");
                return;
            }
            U<PasswordBox>("AdmPw").Password = ""; U<PasswordBox>("AdmPw2").Password = "";
            U("PanelAdminLogin").Visibility = Visibility.Collapsed;
            App.IsAdmin = true;
            App.Log("admin: acceso");
            ApplyLocks();
            OpenAdminPanel();
        }
        public void OpenAdminPanel()
        {
            foreach (string k in LockKeys) U<ToggleButton>("AdmLock" + k).IsChecked = GetLockSet(k);
            U<ToggleButton>("AdmIgnorePause").IsChecked = Admin.IgnorePause;
            U<ToggleButton>("AdmOwnerPC").IsChecked = App.IsOwnerPC;
            U<ToggleButton>("AdmFbSync").IsChecked = Config.Cur.FbSync;
            UpdateAdmFbState();
            U<TextBlock>("AdmLocksMsg").Text = ""; U<TextBlock>("AdmPwMsg").Text = "";
            U<TextBlock>("AdmInfo").Text = string.Format(AdmT("Versión {0} · datos: {1} · programa: {2}"), App.Version, App.DataDir, App.AppDir);
            UpdateRemoteText();
            FillAdminLog();
            U("PanelAdmin").Visibility = Visibility.Visible;
        }

        static readonly Regex LogFatal = new Regex(@"^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  FATAL  (.+)$", RegexOptions.IgnoreCase);
        static readonly Regex LogLine = new Regex(@"^(\d{4})-(\d\d)-(\d\d) (\d\d:\d\d):\d\d  \[([^\]]*)\] (.+)$", RegexOptions.IgnoreCase);
        public List<AdmLogItem> FillAdminLog()
        {
            List<AdmLogItem> items = new List<AdmLogItem>();
            int nAdm = 0, nTeam = 0, nErr = 0;
            bool onlyErr = AdmLogFilter == "errors";
            if (File.Exists(App.LogPath))
            {
                List<string> lines = new List<string>(File.ReadAllLines(App.LogPath, Encoding.UTF8));
                if (lines.Count > 3000) lines.RemoveRange(0, lines.Count - 3000);    // -Tail 3000
                lines.Reverse();
                foreach (string ln0 in lines)
                {
                    string ln = ln0;
                    // los cierres inesperados se guardan con otro formato: «fecha  FATAL  mensaje»
                    Match f = LogFatal.Match(ln);
                    if (f.Success) ln = f.Groups[1].Value + "-" + f.Groups[2].Value + "-" + f.Groups[3].Value + " " + f.Groups[4].Value + ":00  [] fatal: " + f.Groups[5].Value;
                    Match m = LogLine.Match(ln);
                    if (!m.Success) continue;
                    string txt = m.Groups[6].Value, user = m.Groups[5].Value;
                    string time = m.Groups[3].Value + "/" + m.Groups[2].Value + "/" + m.Groups[1].Value + " " + m.Groups[4].Value;
                    bool isErr = TestErrorLogLine(txt);
                    if (isErr) nErr++;
                    if (onlyErr)
                    {
                        if (!isErr || items.Count >= 400) continue;
                        AdmLogItem e = new AdmLogItem();
                        e.Time = time; e.User = IsAdminUser(user) ? AdmT("Admin") : user.Length > 0 ? user : AdmT("Sistema");
                        e.UserBg = "#FDECEC"; e.UserFg = "#B42318"; e.Text = txt; e.Fg = "#B42318";
                        items.Add(e);
                        continue;
                    }
                    if (PsH.Is(txt, "^(Load-Data|Save-FromForm|clic |DEBUG|pestaña no leída|aviso mostrado: |fatal: )")) continue;
                    string kind = IsAdminUser(user) ? "admin" : user.Length > 0 ? "team" : "system";
                    if (kind == "admin") nAdm++; else if (kind == "team") nTeam++;
                    if (AdmLogFilter == "team" && kind != "team") continue;
                    if (AdmLogFilter == "admin" && kind != "admin") continue;
                    string fg = "#2B2526";
                    if (PsH.Is(txt, "fallid|error|no se pudo|incorrect")) fg = "#B42318";
                    else if (PsH.Is(txt, "^admin")) fg = "#9E1820";
                    else if (PsH.Is(txt, "^(añadió|editó)")) fg = "#1E1A1B";
                    else if (PsH.Is(txt, "^(copia externa|sesión)")) fg = "#6E6466";
                    AdmLogItem it = new AdmLogItem();
                    it.Time = time;
                    it.User = kind == "admin" ? AdmT("Admin") : kind == "team" ? user : AdmT("Sistema");
                    it.UserBg = kind == "admin" ? "#9E1820" : kind == "team" ? "#F3EFEF" : "#FFFFFF";
                    it.UserFg = kind == "admin" ? "#FFFFFF" : kind == "team" ? "#2B2526" : "#A69C9E";
                    it.Text = ConvertLogText(txt); it.Fg = fg;
                    items.Add(it);    // como en PowerShell: el tope de 400 solo se mira después de añadir (todo menos errores)
                }
            }
            U<ItemsControl>("AdmLog").ItemsSource = items;
            U<TextBlock>("AdmLogTitle").Text = onlyErr ? string.Format(AdmT("Errores · {0} en este PC (los más recientes arriba)"), nErr) : string.Format(AdmT("Actividad · {0} del equipo · {1} del admin"), nTeam, nAdm);
            U<ContentControl>("AdmLogErr").Content = nErr > 0 ? AdmT("Errores") + " (" + nErr + ")" : AdmT("Errores");
            return items;
        }
        static bool IsAdminUser(string u) { return string.Equals(u, "ADMIN", StringComparison.CurrentCultureIgnoreCase); }    // -eq de PowerShell
        // líneas del registro que son fallos (también los que no enseñan ningún aviso)
        public static bool TestErrorLogLine(string txt)
        {
            return PsH.Is(txt, @"^(fatal: |UI ERROR|aviso mostrado: |Load-Data: |Save-FromForm: |Guardar ERROR|Excel no actualizado|sincronizar con el Excel: |actualizar desde el Excel: |actualización: |pestaña no leída|listas \(|importar \(|config ilegible|accesos directos: |icono: |logo: |logo no válido|recibo: |exportar PDF: |calendario: |copia externa fallida|copia de seguridad fallida|admin: no se pudo|DEBUG (error|editar|conflicto|actualización automática): )");
        }
        void ShowAdminTab(string name) { foreach (string p in new string[] { "AdmPageLocks", "AdmPageLog", "AdmPageTools" }) U(p).Visibility = p == name ? Visibility.Visible : Visibility.Collapsed; }
        public void SaveAdminLocks()
        {
            foreach (string k in LockKeys) Admin.Locks[k] = U<ToggleButton>("AdmLock" + k).IsChecked == true;
            Admin.IgnorePause = U<ToggleButton>("AdmIgnorePause").IsChecked == true;
            Admin.OwnerPC = U<ToggleButton>("AdmOwnerPC").IsChecked == true;
            App.IsOwnerPC = TestOwnerPC();
            UpdateAdmFbState();
            SaveAdmin();
            ApplyLocks();
            List<string> on = new List<string>(); foreach (string k in LockKeys) if (GetLockSet(k)) on.Add(k);
            App.Log("admin: bloqueos guardados (" + string.Join(", ", on) + ")");
            U<TextBlock>("AdmLocksMsg").Text = AdmT("✓ Guardado");
        }
        // estado de la sincronización de este PC, debajo de su casilla
        void UpdateAdmFbState()
        {
            string txt = AdmT("Apagada: este PC no envía ni recibe clientes.");
            if (Config.Cur.FbSync)
            {
                FbState st = null; try { st = FbState.Load(); } catch { }
                int n = st != null ? st.Pend.Count : 0;
                txt = st != null && st.Joined == "yes" ? string.Format(AdmT("Activada · unido · {0} cambio(s) por enviar"), n) : AdmT("Activada · primera sincronización pendiente");    // shortcut: el último error de Firebase llega con la 5b
            }
            U<TextBlock>("AdmFbState").Text = txt;
        }
        public void ChangeAdminPw()
        {
            string a = U<PasswordBox>("AdmNewPw").Password, b = U<PasswordBox>("AdmNewPw2").Password;
            TextBlock m = U<TextBlock>("AdmPwMsg");
            if (a.Length < 6) { m.Foreground = Br("#B42318"); m.Text = AdmT("Mínimo 6 caracteres."); return; }
            if (a != b) { m.Foreground = Br("#B42318"); m.Text = AdmT("Las dos contraseñas no son iguales."); return; }
            Admin.Salt = Guid.NewGuid().ToString("N");
            Admin.Hash = PwHash(a, Admin.Salt);
            SaveAdmin();
            U<PasswordBox>("AdmNewPw").Password = ""; U<PasswordBox>("AdmNewPw2").Password = "";
            m.Foreground = Br("#1B7F45"); m.Text = AdmT("✓ Contraseña cambiada");
            App.Log("admin: contraseña cambiada");
        }
        public void CloseAdmin()
        {
            App.Log("admin: salida");
            App.IsAdmin = false;
            U("PanelAdmin").Visibility = Visibility.Collapsed;
            ApplyLocks();
        }
        public void ApplyPause()
        {
            object j = Remote;
            bool paused = j != null && Truthy(Json.Get(j, "pausa")) && !Admin.IgnorePause && !App.IsAdmin && !App.IsOwnerPC;
            if (paused)
            {
                string motivo = Ps.Str(Json.Get(j, "motivo"));
                U<TextBlock>("PausedText").Text = motivo.Trim().Length > 0 ? motivo : AdmT("El programa está en pausa temporalmente.");
                U("PanelPaused").Visibility = Visibility.Visible;
                App.Log("pausa remota activa");
            }
            else U("PanelPaused").Visibility = Visibility.Collapsed;
        }
        public void UpdateRemoteText()
        {
            object j = Remote;
            string bas = AdmT("GESTEM lee el archivo control.json de tu repositorio de GitHub al abrirse (nunca envía datos). Campos: «mensaje» (texto que verán una vez), «pausa» (true/false) y «motivo».");
            TextBlock t = U<TextBlock>("AdmRemoteTxt");
            if (j == null) { t.Text = bas + "\n\n" + AdmT("Estado: no hay control.json o no se pudo leer → todo normal."); return; }
            string msg = Ps.Str(Json.Get(j, "mensaje"));
            string m = msg.Trim().Length > 0 ? "«" + msg + "»" : AdmT("(vacío)");
            string pz = Truthy(Json.Get(j, "pausa")) ? AdmT("SÍ") : AdmT("no");
            t.Text = bas + "\n\n" + string.Format(AdmT("Estado actual → mensaje: {0}   ·   pausa: {1}"), m, pz);
        }
    }
}

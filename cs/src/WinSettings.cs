// Etapa 4b: Ajustes y empresa. Copiado de Load-SettingsPanel, Load-ImageFile, Get-LogoBrush, Save-CompanyLogo, Choose-Logo,
// Apply-Branding, Import-DefaultBranding, Choose-ExcelFile, Choose-DocsFolder, Set-DocsRoot, Update-ExtBackupInfo,
// Choose-ExtFolder y Save-Settings de GESTEM.ps1 (y sus botones).
// Repair-Shortcuts no se traslada: corrige accesos directos a GESTEM.ps1 (lo hará el instalador de la etapa 7).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Gestem
{
    public partial class MainWin
    {
        public static readonly string[] LangNames = { "Español", "English", "Français", "العربية" };    // $script:Langs (mismo orden que Tx.Langs)
        public string PendingLogo = "";
        BitmapSource appLogo;

        void WireSettings()
        {
            U<Button>("ChangeFileBtn").Click += delegate { ChooseExcelFile(); };
            U<Button>("ChangeDocsBtn").Click += delegate { ChooseDocsFolder(); };
            U<Button>("LogoBtn").Click += delegate { ChooseLogo(); };
            U<Button>("LogoClearBtn").Click += delegate { PendingLogo = ""; U<System.Windows.Shapes.Shape>("LogoPreview").Fill = LogoBrush(""); };
            U<Button>("SaveSettingsBtn").Click += delegate { SaveSettings(); };
            U<Button>("BackupsBtn").Click += delegate { Process.Start("explorer.exe", App.BackupDir); };
            U<Button>("ExtDirBtn").Click += delegate { ChooseExtFolder(); };
            U<Button>("ExtClearBtn").Click += delegate { U<TextBox>("ExtDirBox").Text = ""; UpdateExtBackupInfo(); };
            U<Button>("ExtOpenBtn").Click += delegate
            {
                string d = Ps.Str(U<TextBox>("ExtDirBox").Text).Trim();
                if (d.Length == 0) return;
                try { if (!Directory.Exists(d)) { App.Guard(d, "copia externa"); Directory.CreateDirectory(d); } Process.Start("explorer.exe", "\"" + d + "\""); }
                catch (Exception e) { ShowError(e.Message); }
            };
        }

        public void SetSettingsTexts()
        {
            U<TextBlock>("SetSub").Text = L("setSub");
            U<TextBlock>("ExtDirL").Text = L("extDirL"); U<ContentControl>("ExtDirBtn").Content = L("extChooseBtn"); U<ContentControl>("ExtClearBtn").Content = L("btnLogoClear");
            U<ContentControl>("ExtOneDriveBtn").Content = L("extOneDrive"); U<ContentControl>("ExtNowBtn").Content = L("extNow"); U<ContentControl>("ExtOpenBtn").Content = L("extOpen"); U<TextBlock>("ExtNote").Text = L("extNote");
            foreach (string k in new string[] { "File", "Ext", "Comp", "Team", "Pref", "Help" }) { U<TextBlock>("Sec" + k + "T").Text = L("sec" + k + "T"); U<TextBlock>("Sec" + k + "D").Text = L("sec" + k + "D"); }
            U<TextBlock>("SetCollL").Text = L("setColleagues");
            U<TextBlock>("InvSecL").Text = L("invSec"); U<TextBlock>("InvHintL").Text = L("invHint"); U<TextBlock>("InvNameL").Text = L("invName"); U<TextBlock>("InvActL").Text = L("invAct"); U<TextBlock>("InvNifL").Text = L("invNif");
            U<TextBlock>("TasasL").Text = L("tasasL");
            U<TextBlock>("InvIbanL").Text = L("invIban"); U<TextBlock>("InvDomL").Text = L("invDom"); U<TextBlock>("InvPobL").Text = L("invPob"); U<TextBlock>("InvNextL").Text = L("invNext");
            U<TextBlock>("SetCompanyL").Text = L("setCompany"); U<TextBlock>("SetLogoL").Text = L("setLogo"); U<ContentControl>("LogoBtn").Content = L("btnLogo"); U<ContentControl>("LogoClearBtn").Content = L("btnLogoClear");
            U<TextBlock>("SetTitle").Text = L("setTitle"); U<TextBlock>("SetFileL").Text = L("setFile"); U<ContentControl>("ChangeFileBtn").Content = L("btnChange");
            U<TextBlock>("SetDocsL").Text = L("setDocs"); U<ContentControl>("ChangeDocsBtn").Content = L("btnChange"); U<TextBlock>("DocsNote").Text = L("docsNote");
            U<ContentControl>("ImportXlBtn").Content = L("btnImport"); U<TextBlock>("DataNote").Text = L("dataNote");
            U<TextBlock>("SetLangL").Text = L("setLang"); U<TextBlock>("SetNatL").Text = L("setNat"); U<ContentControl>("BackupChk").Content = L("setBackup");
            U<ContentControl>("OpenAfterChk").Content = L("setOpen"); U<TextBlock>("SetNote").Text = L("setNote");
            U<ContentControl>("SaveSettingsBtn").Content = L("btnSaveSettings"); U<ContentControl>("BackupsBtn").Content = L("btnBackups"); U<ContentControl>("UpdateBtn").Content = L("btnUpdate");
            U<TextBlock>("WelcomeText").Text = L("welcome");
            U<TextBlock>("AboutText").Text = L("aboutText"); U<TextBlock>("AboutBy").Text = L("aboutBy"); U<TextBlock>("AboutVer").Text = string.Format(L("aboutVer"), App.Version);
        }

        public void LoadSettingsPanel()
        {
            U<TextBox>("FilePathBox").Text = Config.Cur.ExcelPath;
            U<TextBox>("DocsDirBox").Text = Ps.Str(Docs.Root());
            ComboBox lb = U<ComboBox>("LangBox");
            lb.Items.Clear();
            foreach (string n in LangNames) lb.Items.Add(n);
            lb.SelectedIndex = Array.IndexOf(Tx.Langs, Config.Cur.Language);
            ComboBox nb = U<ComboBox>("NatBox");
            nb.Items.Clear();
            foreach (string n in Store.Nacionalidades) nb.Items.Add(n);
            nb.Text = Config.Cur.DefaultNationality;
            U<TextBox>("ColleaguesBox").Text = string.Join(", ", Config.Cur.Colleagues);
            U<TextBox>("CompanyBox").Text = Ps.Str(Config.Cur.CompanyName);
            U<TextBox>("InvNameBox").Text = Ps.Str(Config.Cur.InvName); U<TextBox>("InvActBox").Text = Ps.Str(Config.Cur.InvActivity); U<TextBox>("InvNifBox").Text = Ps.Str(Config.Cur.InvNif);
            U<TextBox>("InvIbanBox").Text = Ps.Str(Config.Cur.InvIban); U<TextBox>("InvDomBox").Text = Ps.Str(Config.Cur.InvAddress); U<TextBox>("InvPobBox").Text = Ps.Str(Config.Cur.InvTown);
            string nx; try { nx = Invoices.Get().Next.ToString("000"); } catch { nx = ""; }
            U<TextBox>("InvNextBox").Text = nx;
            U<TextBox>("TasasBox").Text = Ps.Str(Config.Cur.TasasCanje);
            PendingLogo = Ps.Str(Config.Cur.CompanyLogo);
            U<System.Windows.Shapes.Shape>("LogoPreview").Fill = LogoBrush(PendingLogo);
            U<TextBox>("ExtDirBox").Text = Ps.Str(Config.Cur.ExtBackupDir);
            string od = Environment.GetEnvironmentVariable("OneDrive"); if (string.IsNullOrEmpty(od)) od = Environment.GetEnvironmentVariable("OneDriveConsumer");
            U("ExtOneDriveBtn").Visibility = !string.IsNullOrEmpty(od) && (Directory.Exists(od) || File.Exists(od)) ? Visibility.Visible : Visibility.Collapsed;
            UpdateExtBackupInfo();
            U<CheckBox>("BackupChk").IsChecked = Config.Cur.Backup;
            U<CheckBox>("OpenAfterChk").IsChecked = Config.Cur.OpenAfterSave;
            U("WelcomeText").Visibility = string.IsNullOrEmpty(Config.Cur.ExcelPath) ? Visibility.Visible : Visibility.Collapsed;
        }

        // --- empresa: nombre y logo ---
        static BitmapImage LoadImageFile(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return null;
            try
            {
                BitmapImage bmp = new BitmapImage();
                bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad; bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bmp.UriSource = new Uri(path); bmp.DecodePixelWidth = 256; bmp.EndInit(); bmp.Freeze();
                return bmp;
            }
            catch (Exception e) { App.Log("logo no válido: " + e.Message); return null; }
        }
        ImageBrush LogoBrush(string path)
        {
            ImageSource src = LoadImageFile(path);
            if (src == null) src = appLogo;
            if (src == null) return null;
            ImageBrush b = new ImageBrush(src); b.Stretch = Stretch.Uniform;
            return b;
        }
        // copia la imagen elegida dentro de la carpeta de datos (así no depende del archivo original)
        static string SaveCompanyLogo(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return "";
            string ext = Path.GetExtension(path).ToLower();
            string dest = Path.Combine(App.DataDir, "empresa" + ext);
            if (string.Equals(Path.GetFullPath(path), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase)) return dest;
            try
            {
                App.Guard(dest, "logo");
                foreach (FileInfo f in new DirectoryInfo(App.DataDir).GetFiles("empresa.*"))
                    if (Array.IndexOf(new string[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif" }, f.Extension.ToLower()) >= 0) try { f.Delete(); } catch { }
                File.Copy(path, dest, true);
                return dest;
            }
            catch (Exception e) { App.Log("logo: " + e.Message); return path; }
        }
        void ChooseLogo()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Title = L("chooseLogo");
            dlg.Filter = "Imagen (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp";
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            if (dlg.ShowDialog(Win) == true)
            {
                if (LoadImageFile(dlg.FileName) == null) { ShowError(L("chooseLogo")); return; }
                PendingLogo = dlg.FileName;
                U<System.Windows.Shapes.Shape>("LogoPreview").Fill = LogoBrush(PendingLogo);
            }
        }
        public void ApplyBranding()
        {
            string name = Ps.Str(Config.Cur.CompanyName).Trim();
            ImageBrush brush = LogoBrush(Ps.Str(Config.Cur.CompanyLogo));
            U<System.Windows.Shapes.Shape>("SideLogo").Fill = brush;
            U<System.Windows.Shapes.Shape>("SessLogo").Fill = brush;
            U<TextBlock>("SideCompany").Text = name.ToUpper();
            U("SideCompany").Visibility = name.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
            U<TextBlock>("SessCompany").Text = name.Length > 0 ? name : "GESTEM";
            U<TextBlock>("AboutCompany").Text = name;
            Win.Title = name.Length > 0 ? "GESTEM MANAGER · " + name : "GESTEM MANAGER";
            if (!App.Real) Win.Title += "  —  " + TestMode.Banner;
        }
        // primera vez: si el instalador trae empresa.png / empresa.txt junto al programa, se usan solos
        static void ImportDefaultBranding()
        {
            bool changed = false;
            if (string.IsNullOrEmpty(Config.Cur.CompanyLogo))
                foreach (string e in new string[] { ".png", ".jpg", ".jpeg" })
                {
                    string f = Path.Combine(App.AppDir, "empresa" + e);
                    if (File.Exists(f)) { Config.Cur.CompanyLogo = SaveCompanyLogo(f); changed = true; break; }
                }
            if (string.IsNullOrEmpty(Config.Cur.CompanyName))
            {
                string t = Path.Combine(App.AppDir, "empresa.txt");
                if (File.Exists(t)) { using (StreamReader r = new StreamReader(t, System.Text.Encoding.UTF8)) Config.Cur.CompanyName = Ps.Str(r.ReadLine()).Trim(); changed = true; }
            }
            if (changed) Config.Save();
        }
        // icono de GESTEM: ventana, «Acerca de» y logo por defecto
        void InitBranding()
        {
            try
            {
                using (Stream s = typeof(MainWin).Assembly.GetManifestResourceStream("gestem.ico"))
                {
                    BitmapDecoder dec = BitmapDecoder.Create(s, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                    foreach (BitmapFrame f in dec.Frames) if (appLogo == null || f.PixelWidth > appLogo.PixelWidth) appLogo = f;
                }
                Win.Icon = appLogo;
            }
            catch (Exception e) { App.Log("icono: " + e.Message); }
            U<Image>("AboutImg").Source = appLogo;
            PendingLogo = "";
            ImportDefaultBranding();
            ApplyBranding();
        }

        // --- archivo, documentos, copia externa ---
        bool ChooseExcelFile()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Title = L("chooseFile");
            dlg.Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
            if (!string.IsNullOrEmpty(Config.Cur.ExcelPath)) { try { dlg.InitialDirectory = Path.GetDirectoryName(Config.Cur.ExcelPath); } catch { } }
            else dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (dlg.ShowDialog(Win) == true) { U<TextBox>("FilePathBox").Text = dlg.FileName; return true; }
            return false;
        }
        static string PickFolder(string desc, string start)
        {
            using (System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog())
            {
                dlg.Description = desc; dlg.ShowNewFolderButton = true;
                if (!string.IsNullOrEmpty(start)) dlg.SelectedPath = start;
                return dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK ? dlg.SelectedPath : null;
            }
        }
        void ChooseDocsFolder()
        {
            string p = PickFolder(L("docsChoose"), U<TextBox>("DocsDirBox").Text);
            if (p == null) return;
            if (Path.GetFileName(p) != "GESTEM documentos") p = Path.Combine(p, "GESTEM documentos");    // con este nombre los otros PCs la encuentran solos
            U<TextBox>("DocsDirBox").Text = p;
        }
        // cambia la carpeta de documentos llevando a la nueva los que había: se mueven si estaban en la de este PC (al lado del Excel);
        // de otra carpeta (quizá compartida) solo se copian
        public static MergeResult SetDocsRoot(string nw)
        {
            string old = Docs.Root();
            MergeResult r = Docs.Merge(old, nw, Ps.Str(old) != Ps.Str(Docs.LocalRoot()));
            Config.Cur.DocsDir = nw; Config.Save();
            Docs.UpdateMap();
            return r;
        }
        void UpdateExtBackupInfo()
        {
            string d = Ps.Str(U<TextBox>("ExtDirBox").Text).Trim();
            string last = Ps.Str(Config.Cur.LastExtBackup), when = last;
            DateTime w; if (DateTime.TryParseExact(last, "yyyy-MM-dd HH:mm", null, System.Globalization.DateTimeStyles.None, out w)) when = w.ToString("dd/MM/yyyy HH:mm");
            TextBlock t = U<TextBlock>("ExtLastTxt");
            t.Text = last.Length > 0 && d.Length > 0 && d == Ps.Str(Config.Cur.ExtBackupDir) ? string.Format(L("extLast"), when) : d.Length > 0 ? L("extNever") : L("extNone");
            t.Foreground = Br(last.Length > 0 && d.Length > 0 ? "#1B7F45" : "#B42318");
            U<Button>("ExtNowBtn").IsEnabled = d.Length > 0;
            U<Button>("ExtOpenBtn").IsEnabled = d.Length > 0;
        }
        void ChooseExtFolder()
        {
            string p = PickFolder(L("extChoose"), U<TextBox>("ExtDirBox").Text);
            if (p == null) return;
            U<TextBox>("ExtDirBox").Text = Path.Combine(p, "GESTEM copias");
            UpdateExtBackupInfo();
        }

        public void SaveSettings()
        {
            Config c = Config.Cur;
            string oldPath = c.ExcelPath;
            c.ExcelPath = U<TextBox>("FilePathBox").Text;
            int li = U<ComboBox>("LangBox").SelectedIndex; if (li >= 0) c.Language = Tx.Langs[li];
            string nat = Ps.Str(U<ComboBox>("NatBox").Text).Trim().ToUpper(); if (nat.Length > 0) c.DefaultNationality = nat;
            c.CompanyName = Ps.Str(U<TextBox>("CompanyBox").Text).Trim();
            c.InvName = Ps.Str(U<TextBox>("InvNameBox").Text).Trim(); c.InvActivity = Ps.Str(U<TextBox>("InvActBox").Text).Trim(); c.InvNif = Ps.Str(U<TextBox>("InvNifBox").Text).Trim().ToUpper();
            c.InvIban = Ps.Str(U<TextBox>("InvIbanBox").Text).Trim().ToUpper(); c.InvAddress = Ps.Str(U<TextBox>("InvDomBox").Text).Trim(); c.InvTown = Ps.Str(U<TextBox>("InvPobBox").Text).Trim();
            c.TasasCanje = PsH.Rep(Ps.Str(U<TextBox>("TasasBox").Text).Trim(), @"\s*[;\r\n]+\s*", "; ");
            int nx; if (int.TryParse(Ps.Str(U<TextBox>("InvNextBox").Text).Trim(), out nx) && nx > 0) { try { Invoices.SetNext(nx); } catch (Exception e) { App.Log("facturas: " + e.Message); } }
            c.CompanyLogo = SaveCompanyLogo(PendingLogo);
            ApplyBranding();
            List<string> col = new List<string>();
            foreach (string x in Ps.Str(U<TextBox>("ColleaguesBox").Text).Split(',')) { string y = x.Trim().ToUpper(); if (y.Length > 0) col.Add(y); }
            c.Colleagues = col;
            G.Lists["colleagues"] = col.ToArray();
            string newExt = Ps.Str(U<TextBox>("ExtDirBox").Text).Trim();
            bool extChanged = newExt != Ps.Str(c.ExtBackupDir);
            c.ExtBackupDir = newExt;
            if (extChanged) { c.LastExtBackup = ""; if (newExt.Length > 0) BackupExternal(false); }
            c.Backup = U<CheckBox>("BackupChk").IsChecked == true;
            c.OpenAfterSave = U<CheckBox>("OpenAfterChk").IsChecked == true;
            string newDocs = Ps.Str(U<TextBox>("DocsDirBox").Text).Trim();
            if (newDocs.Length > 0 && newDocs != Ps.Str(Docs.Root()))
            {
                try { SetDocsRoot(newDocs); App.Log("documentos: carpeta cambiada a " + newDocs); }
                catch (Exception e) { App.Log("documentos: " + e.Message); ShowError(e.Message); }
            }
            Config.Save();
            if (c.ExcelPath != oldPath) Data.LoadData(true, false);
            ApplyLanguage();
            LoadSettingsPanel();
            App.Log("cambió los ajustes");
            SetStatus(L("settingsSaved"));
        }
    }
}

// Entrar con contraseña (solo GESTEM.exe): lista de usuarios con su icono; la primera vez (o tras restablecerla) cada persona crea su
// contraseña (mín. 4, repetida) y elige icono (color con iniciales o foto). Admin: «Restablecer contraseña». Ver Profiles.cs.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;

namespace Gestem
{
    public partial class MainWin
    {
        string authUser = "", pickColor = "", pickPhoto = "";    // pickPhoto: nombre de la foto guardada, o ruta completa de una recién elegida
        bool authCreate;
        public static Func<string> PhotoPickHook;    // pruebas: la foto elegida sin el diálogo
        Task<List<Profile>> profTask;

        void WireLogin()
        {
            U<Button>("SessEnter").Click += delegate { AuthSubmit(); };
            U<Button>("SessBack").Click += delegate { ShowSession(); };
            foreach (string n in new string[] { "SessPw", "SessPw2" })
                U<PasswordBox>(n).KeyDown += delegate(object s, KeyEventArgs e) { if (e.Key == Key.Return) { AuthSubmit(); e.Handled = true; } };
            U<Button>("SessPhoto").Click += delegate { PickPhoto(); };
            U<Button>("SessNoPhoto").Click += delegate { pickPhoto = ""; RefreshAuthIcon(); };
            U<ToggleButton>("AdmTabUsers").Checked += delegate { FillAdminUsers(); ShowAdminTab("AdmPageUsers"); };
        }

        static Brush Hex(string c) { return (Brush)new BrushConverter().ConvertFromString(c); }
        static BitmapImage LoadImage(string path, int px)
        {
            BitmapImage b = new BitmapImage();
            b.BeginInit(); b.CacheOption = BitmapCacheOption.OnLoad; b.CreateOptions = BitmapCreateOptions.IgnoreImageCache;    // sin dejar el archivo abierto
            if (px > 0) b.DecodePixelWidth = px;
            b.UriSource = new Uri(path); b.EndInit(); b.Freeze();
            return b;
        }

        // icono: la foto, o un círculo de color con las iniciales
        public static FrameworkElement Avatar(string name, double size, string color = null, string photoPath = null)
        {
            Grid g = new Grid(); g.Width = size; g.Height = size;
            string photo = photoPath ?? Profiles.PhotoPath(name);
            Ellipse e = new Ellipse();
            if (photo != null && File.Exists(photo))
            {
                try { ImageBrush ib = new ImageBrush(LoadImage(photo, 0)); ib.Stretch = Stretch.UniformToFill; e.Fill = ib; g.Children.Add(e); g.Tag = "foto"; return g; } catch { }
            }
            e.Fill = Hex(color ?? Profiles.ColorOf(name)); g.Children.Add(e);
            TextBlock t = new TextBlock(); t.Text = Profiles.Initials(name); t.Foreground = Brushes.White; t.FontWeight = FontWeights.SemiBold; t.FontSize = size * 0.38;
            t.HorizontalAlignment = HorizontalAlignment.Center; t.VerticalAlignment = VerticalAlignment.Center;
            g.Children.Add(t); g.Tag = "color";
            return g;
        }

        // nombres de la lista: los de siempre (Colaboradores/compañeras) y los que tienen perfil (también los llegados de otro PC)
        static List<string> LoginNames(List<string> people)
        {
            List<string> l = new List<string>(people);
            foreach (string n in Profiles.Names()) if (!PsH.ContainsI(l, n)) l.Add(n);
            return l;
        }

        // al abrir: la última persona va directa a su contraseña; «Cambiar de usuario» vuelve a la lista
        public void ShowLogin()
        {
            ShowSession();
            string last = Ps.Str(Config.Cur.LastUser).Trim();
            if (last.Length > 0) ShowAuth(last);
        }

        public void ShowAuth(string who)
        {
            who = Ps.Str(who).Trim().ToUpper();
            if (who.Length == 0) return;
            authUser = who; authCreate = !Profiles.HasPassword(who);
            Profile p = Profiles.Get(who);
            pickColor = Profiles.ColorOf(who); pickPhoto = p != null ? p.Photo : "";
            U<TextBlock>("SessWho").Text = who;
            U<TextBlock>("SessAuthHint").Text = authCreate ? L("pwNewHint") : L("loginHint");
            U<TextBlock>("SessIconLbl").Text = L("iconPick");
            U<ContentControl>("SessPhoto").Content = L("iconPhoto"); U<ContentControl>("SessNoPhoto").Content = L("iconNoPhoto");
            U<TextBlock>("SessPwLbl").Text = authCreate ? L("pwNew") : L("loginPw");
            U<TextBlock>("SessPw2Lbl").Text = L("pwRepeat");
            U<ContentControl>("SessEnter").Content = authCreate ? L("pwSave") : L("loginEnter");
            U<ContentControl>("SessBack").Content = L("loginOther");
            Visibility cv = authCreate ? Visibility.Visible : Visibility.Collapsed;
            U("SessIconBox").Visibility = cv; U("SessPw2").Visibility = cv; U("SessPw2Lbl").Visibility = cv;
            U<PasswordBox>("SessPw").Clear(); U<PasswordBox>("SessPw2").Clear(); U<TextBlock>("SessMsg").Text = "";
            BuildSwatches(); RefreshAuthIcon();
            U("SessPick").Visibility = Visibility.Collapsed; U("SessAuth").Visibility = Visibility.Visible;
            U("PanelSession").Visibility = Visibility.Visible;
            Win.Dispatcher.BeginInvoke(new Action(delegate { U<PasswordBox>("SessPw").Focus(); }), System.Windows.Threading.DispatcherPriority.Input);
        }

        void BuildSwatches()
        {
            Panel sw = U<Panel>("SessColors"); sw.Children.Clear();
            foreach (string c in Profiles.Colors)
            {
                Border b = new Border(); b.Width = 30; b.Height = 30; b.CornerRadius = new CornerRadius(15); b.Margin = new Thickness(0, 0, 8, 0); b.Cursor = Cursors.Hand; b.Tag = c;
                b.BorderThickness = new Thickness(c == pickColor ? 3 : 0); b.BorderBrush = Hex("#1E1A1B"); b.Background = Hex(c);
                b.MouseLeftButtonUp += delegate { PickColor((string)b.Tag); };
                sw.Children.Add(b);
            }
        }
        public void PickColor(string c) { pickColor = c; pickPhoto = ""; BuildSwatches(); RefreshAuthIcon(); }

        void RefreshAuthIcon()
        {
            Grid host = U<Grid>("SessAvatar"); host.Children.Clear();
            string ph = pickPhoto.Length == 0 ? "" : System.IO.Path.IsPathRooted(pickPhoto) ? pickPhoto : System.IO.Path.Combine(Profiles.PhotoDir, pickPhoto);
            host.Children.Add(Avatar(authUser, 64, pickColor, ph));
            U("SessNoPhoto").Visibility = pickPhoto.Length > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        void PickPhoto()
        {
            string f = null;
            if (PhotoPickHook != null) f = PhotoPickHook();
            else
            {
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "Fotos (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";
                if (dlg.ShowDialog(Win) == true) f = dlg.FileName;
            }
            if (string.IsNullOrEmpty(f)) return;
            try { LoadImage(f, 64); } catch { ShowError(L("errGeneric")); return; }    // no es una imagen
            pickPhoto = f; RefreshAuthIcon();
        }

        // la foto elegida: cuadrada (centro), 128 px, PNG en %APPDATA%\GESTEM\fotos (solo en este PC)
        static string SavePhoto(string who, string pick)
        {
            if (pick.Length == 0 || !System.IO.Path.IsPathRooted(pick)) return pick;
            BitmapImage src = LoadImage(pick, 0);
            int side = Math.Min(src.PixelWidth, src.PixelHeight);
            BitmapSource sq = new CroppedBitmap(src, new Int32Rect((src.PixelWidth - side) / 2, (src.PixelHeight - side) / 2, side, side));
            double k = 128.0 / side;
            BitmapSource small = new TransformedBitmap(sq, new ScaleTransform(k, k));
            string name = Docs.SafeName(who) + ".png", path = System.IO.Path.Combine(Profiles.PhotoDir, name);
            App.Guard(path, "foto de usuario"); Directory.CreateDirectory(Profiles.PhotoDir);
            PngBitmapEncoder enc = new PngBitmapEncoder(); enc.Frames.Add(BitmapFrame.Create(small));
            using (FileStream fs = File.Create(path)) enc.Save(fs);
            return name;
        }

        public void AuthSubmit()
        {
            PasswordBox a = U<PasswordBox>("SessPw"), b = U<PasswordBox>("SessPw2"); TextBlock msg = U<TextBlock>("SessMsg");
            if (authUser.Length == 0) return;
            if (!authCreate)
            {
                if (Profiles.Check(authUser, a.Password)) { a.Clear(); StartSession(authUser); return; }
                App.Log("contraseña incorrecta para " + authUser);
                msg.Text = L("loginBad"); a.Clear(); a.Focus();
                return;
            }
            string why = Profiles.Validate(a.Password, b.Password);
            if (why.Length > 0) { msg.Text = L(why); return; }
            try
            {
                Profiles.SetIcon(authUser, pickColor, SavePhoto(authUser, pickPhoto));
                Profiles.SetPassword(authUser, a.Password);
            }
            catch (Exception e) { App.LogError("crear contraseña", e); msg.Text = L("errGeneric") + " " + e.Message; return; }
            a.Clear(); b.Clear();
            SyncProfiles(false);
            StartSession(authUser);
        }

        // ---------- administrador ----------
        public void FillAdminUsers()
        {
            Panel list = U<Panel>("AdmUsersList"); list.Children.Clear();
            List<string> names = LoginNames(People());
            if (names.Count == 0) { TextBlock t = new TextBlock(); t.Text = AdmT("Todavía no hay usuarios."); t.Foreground = Hex("#6E6466"); list.Children.Add(t); return; }
            foreach (string n in names)
            {
                string who = n.Trim().ToUpper();
                DockPanel row = new DockPanel(); row.Margin = new Thickness(0, 0, 0, 10);
                Button rb = new Button(); rb.Style = (Style)Win.FindResource("Btn"); rb.Content = AdmT("Restablecer contraseña"); rb.Height = 34; rb.Padding = new Thickness(14, 0, 14, 0);
                rb.IsEnabled = Profiles.HasPassword(who); rb.Tag = who;
                rb.Click += delegate { AdminResetPassword((string)rb.Tag); };
                DockPanel.SetDock(rb, Dock.Right); row.Children.Add(rb);
                FrameworkElement av = Avatar(who, 34); av.Margin = new Thickness(0, 0, 12, 0); DockPanel.SetDock(av, Dock.Left); row.Children.Add(av);
                StackPanel tx = new StackPanel(); tx.VerticalAlignment = VerticalAlignment.Center;
                TextBlock nm = new TextBlock(); nm.Text = who; nm.FontWeight = FontWeights.SemiBold; nm.Foreground = Hex("#1E1A1B"); tx.Children.Add(nm);
                TextBlock st = new TextBlock(); st.Text = AdmT(Profiles.HasPassword(who) ? "tiene contraseña" : "creará una contraseña al entrar"); st.Foreground = Hex("#8A7F81"); st.FontSize = 12.5; tx.Children.Add(st);
                row.Children.Add(tx);
                list.Children.Add(row);
            }
        }
        public void AdminResetPassword(string who)
        {
            Profiles.Reset(who, "admin");
            SyncProfiles(false);
            FillAdminUsers();
            U<TextBlock>("AdmUsersMsg").Text = string.Format(AdmT("Contraseña de {0} restablecida: creará una nueva al entrar."), who);
        }

        // ---------- sincronización con Firebase (solo hash, sal y estado) ----------
        // wait: al abrir, antes de mirar control.json (así un restablecimiento ya hecho en otro PC no se repite)
        public void SyncProfiles(bool wait)
        {
            string k = "";
            if (Profiles.RemoteHook == null)
            {
                if (!Config.Cur.FbSync || !(App.Real || Fb.TestRun)) return;
                k = Updater.GetKey(); if (k.Length == 0) return;
            }
            if (profTask != null && !profTask.IsCompleted) return;
            List<Profile> snap = Profiles.Snapshot();
            profTask = Task.Run(delegate { FbAuth a = null; return Profiles.SyncOnce(k, ref a, snap); });
            Action<Task<List<Profile>>> done = delegate(Task<List<Profile>> t)
            {
                if (t.Status != TaskStatus.RanToCompletion) { App.Log("usuarios (Firebase): " + (t.Exception != null ? t.Exception.GetBaseException().Message : "sin respuesta")); return; }
                if (Profiles.Merge(t.Result)) OnProfilesChanged();
            };
            if (wait)
            {
                try { profTask.Wait(15000); } catch { }
                if (profTask.IsCompleted) done(profTask); else App.Log("usuarios (Firebase): sin respuesta");
            }
            else profTask.ContinueWith(done, TaskScheduler.FromCurrentSynchronizationContext());
        }

        void OnProfilesChanged()
        {
            if (U("PanelSession").Visibility == Visibility.Visible)
            {
                if (U("SessAuth").Visibility == Visibility.Visible) { if (authCreate == Profiles.HasPassword(authUser)) ShowAuth(authUser); }    // cambió: ahora tiene (o ya no tiene) contraseña
                else ShowSession();
            }
            if (U("AdmPageUsers").Visibility == Visibility.Visible) FillAdminUsers();
        }
    }
}

// Etapa 3d: panel de documentos del cliente. Copiado de Open-DocsPanel, Fill-DocsList, Add-ClientDocs, Add-DocsFromDialog,
// Open-SelectedDoc, Remove-SelectedDoc y Open-DocsFolder de GESTEM.ps1 (y sus eventos: botones, doble clic, arrastrar archivos).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Gestem
{
    public class DocItem { public string Name { get; set; } public string Info { get; set; } public string Glyph { get; set; } public string Path { get; set; } }

    public partial class MainWin
    {
        public ClientRow DocsRow;    // $script:DocsRow

        void WireDocs()
        {
            U<Button>("DocsBtn").Click += delegate { ClientRow r = SelectedClient(); if (r != null) OpenDocsPanel(r); };
            U<MenuItem>("CmDocs").Click += delegate { ClientRow r = SelectedClient(); if (r != null) OpenDocsPanel(r); };
            cmFolder.Click += delegate { ClientRow r = SelectedClient(); if (r != null) { DocsRow = r; OpenDocsFolder(); } };
            U<Button>("DocsClose").Click += delegate { U("PanelDocs").Visibility = Visibility.Collapsed; };
            U<Button>("DocsAddBtn").Click += delegate { AddDocsFromDialog(); };
            U<Button>("DocsOpenBtn").Click += delegate { if (U<ListBox>("DocsList").SelectedItem != null) OpenSelectedDoc(); else Info(L("docsPick")); };
            U<Button>("DocsRemoveBtn").Click += delegate { RemoveSelectedDoc(true); };
            U<Button>("DocsFolderBtn").Click += delegate { OpenDocsFolder(); };
            U<ListBox>("DocsList").MouseDoubleClick += delegate { OpenSelectedDoc(); };
            foreach (UIElement dz in new UIElement[] { U("PanelDocs"), U("DocsCard") })
            {
                dz.DragEnter += delegate(object s, DragEventArgs e)
                {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop)) { e.Effects = DragDropEffects.Copy; U("DocsDrop").Visibility = Visibility.Visible; } else e.Effects = DragDropEffects.None;
                    e.Handled = true;
                };
                dz.DragOver += delegate(object s, DragEventArgs e) { e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None; e.Handled = true; };
                dz.DragLeave += delegate { U("DocsDrop").Visibility = Visibility.Collapsed; };
                dz.Drop += delegate(object s, DragEventArgs e)
                {
                    U("DocsDrop").Visibility = Visibility.Collapsed;
                    if (e.Data.GetDataPresent(DataFormats.FileDrop)) AddClientDocs((string[])e.Data.GetData(DataFormats.FileDrop));
                    e.Handled = true;
                };
            }
            // clic en el clip de la columna «Docs»: abre directamente los documentos del cliente
            DataGrid g = U<DataGrid>("ClientsGrid");
            g.PreviewMouseLeftButtonUp += delegate(object s, MouseButtonEventArgs e)
            {
                try
                {
                    DependencyObject o = e.OriginalSource as DependencyObject;
                    while (o != null && !(o is DataGridCell))
                    {
                        if (o is System.Windows.Controls.Primitives.DataGridColumnHeader) return;
                        o = (o is Visual || o is System.Windows.Media.Media3D.Visual3D) ? VisualTreeHelper.GetParent(o) : LogicalTreeHelper.GetParent(o);
                    }
                    DataGridCell cell = o as DataGridCell;
                    if (cell != null && cell.Column == U<DataGridColumn>("DocsCol") && cell.DataContext is ClientRow)
                    {
                        g.SelectedItem = cell.DataContext;
                        OpenDocsPanel((ClientRow)cell.DataContext);
                    }
                }
                catch (Exception x) { App.Log("docs clic: " + x.Message); }
            };
        }

        public void SetDocsTexts()
        {
            U<ContentControl>("DocsAddBtn").Content = L("docsAdd"); U<ContentControl>("DocsOpenBtn").Content = L("docsOpen");
            U<ContentControl>("DocsRemoveBtn").Content = L("docsRemove"); U<ContentControl>("DocsFolderBtn").Content = L("docsFolder");
            U<TextBlock>("DocsEmptyT").Text = L("docsEmpty"); U<TextBlock>("DocsDropT").Text = L("docsDrop"); U<FrameworkElement>("DocsClose").ToolTip = L("docsClose");
        }

        void Info(string msg) { if (!App.Quiet) MessageBox.Show(Win, msg, App.Name, MessageBoxButton.OK, MessageBoxImage.Information); }

        public void OpenDocsPanel(ClientRow r)
        {
            if (r == null) return;
            if (Docs.Root() == null) { ShowError(L("noFile")); return; }
            DocsRow = r;
            U<TextBlock>("DocsTitle").Text = L("docsTitle");
            U<TextBlock>("DocsSub").Text = string.Format(L("docsSub"), r.Id, r.Name);
            U<TextBlock>("DocsMsg").Text = "";
            FillDocsList();
            U("PanelDocs").Visibility = Visibility.Visible;
        }

        public List<DocItem> FillDocsList()
        {
            List<DocItem> items = new List<DocItem>();
            string dir = DocsRow == null ? null : Docs.ClientDir(G.T, DocsRow, false);
            if (dir != null && Directory.Exists(dir))
            {
                List<FileInfo> files = Docs.VisibleFiles(dir);
                // Sort-Object LastWriteTime -Descending (estable: a igual hora, el orden de la carpeta)
                List<KeyValuePair<int, FileInfo>> tmp = new List<KeyValuePair<int, FileInfo>>();
                for (int i = 0; i < files.Count; i++) tmp.Add(new KeyValuePair<int, FileInfo>(i, files[i]));
                tmp.Sort(delegate(KeyValuePair<int, FileInfo> a, KeyValuePair<int, FileInfo> b) { int c = b.Value.LastWriteTime.CompareTo(a.Value.LastWriteTime); return c != 0 ? c : a.Key.CompareTo(b.Key); });
                foreach (KeyValuePair<int, FileInfo> kv in tmp)
                {
                    FileInfo f = kv.Value;
                    string ext = f.Extension.ToLowerInvariant();
                    string glyph = ((char)0xE8A5).ToString();
                    if (Array.IndexOf(new string[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".heic", ".webp" }, ext) >= 0) glyph = ((char)0xEB9F).ToString();
                    else if (ext == ".pdf") glyph = ((char)0xEA90).ToString();
                    double kb = Math.Max(1.0, Math.Round(f.Length / 1024.0));
                    string size = kb >= 1024 ? (kb / 1024).ToString("0.0", Ps.ES) + " MB" : kb.ToString(Ps.Inv) + " KB";
                    DocItem it = new DocItem(); it.Name = f.Name; it.Info = size + "  ·  " + f.LastWriteTime.ToString("dd/MM/yyyy HH:mm"); it.Glyph = glyph; it.Path = f.FullName;
                    items.Add(it);
                }
            }
            U<ListBox>("DocsList").ItemsSource = items;
            U("DocsEmpty").Visibility = items.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
            U<TextBlock>("DocsCount").Text = items.Count > 0 ? string.Format(L("docsCount"), items.Count) : "";
            return items;
        }

        public void AddClientDocs(string[] paths)
        {
            ClientRow r = DocsRow;
            if (r == null || paths == null) return;
            List<string> files = new List<string>(); foreach (string p in paths) if (!string.IsNullOrEmpty(p) && File.Exists(p)) files.Add(p);
            if (files.Count == 0) return;
            try
            {
                string dir = Docs.ClientDir(G.T, r, true);
                int n = 0;
                foreach (string p in files)
                {
                    string name = Path.GetFileNameWithoutExtension(p), ext = Path.GetExtension(p);
                    string dst = Path.Combine(dir, name + ext); int k = 2;
                    while (File.Exists(dst)) { dst = Path.Combine(dir, name + " (" + k + ")" + ext); k++; }
                    File.Copy(p, dst);
                    n++;
                }
                App.Log(string.Format("añadió {0} documento(s) al expediente {1}", n, r.Id));
                U<TextBlock>("DocsMsg").Text = string.Format(L("docsAdded"), n);
            }
            catch (Exception e) { ShowError(L("errGeneric") + Environment.NewLine + e.Message); }
            FillDocsList();
            Docs.UpdateMap(); RefreshClientGrid();
        }

        void AddDocsFromDialog()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Title = L("docsAdd");
            dlg.Filter = L("docsFilter") + "|*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.heic;*.webp;*.doc;*.docx;*.xls;*.xlsx|" + L("docsAll") + "|*.*";
            if (dlg.ShowDialog(Win) == true) AddClientDocs(dlg.FileNames);
        }

        void OpenSelectedDoc()
        {
            DocItem it = U<ListBox>("DocsList").SelectedItem as DocItem;
            if (it == null) return;
            try { Process.Start(it.Path); } catch (Exception e) { ShowError(e.Message); }
        }

        // a la papelera (se puede recuperar), como en PowerShell
        public void RemoveSelectedDoc(bool ask)
        {
            DocItem it = U<ListBox>("DocsList").SelectedItem as DocItem;
            if (it == null) { Info(L("docsPick")); return; }
            if (ask && MessageBox.Show(Win, string.Format(L("docsRemoveAsk"), it.Name), App.Name, MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            try
            {
                App.Guard(it.Path, "documentos");
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(it.Path, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                App.Log(string.Format("quitó un documento del expediente {0}", DocsRow.Id));
            }
            catch (Exception e) { ShowError(e.Message); }
            FillDocsList();
            Docs.UpdateMap(); RefreshClientGrid();
        }

        void OpenDocsFolder()
        {
            string dir = Docs.ClientDir(G.T, DocsRow, true);
            Process.Start("explorer.exe", "\"" + dir + "\"");
        }
    }
}

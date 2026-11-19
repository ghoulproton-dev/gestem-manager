// Exportar la lista de clientes (botón «Exportar ▾»: PDF o CSV). Copiado de Get-ExportFile, Export-ClientsCsv, Draw-ClientList y
// Export-ClientsPdf de GESTEM.ps1: los clientes que se ven en la lista (con búsqueda y filtros), PDF apaisado de 25 por página.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace Gestem
{
    public partial class MainWin
    {
        public static Func<string, string, string> ExportFileHook;    // pruebas: ruta del archivo en vez del diálogo «Guardar como»

        void WireExport()
        {
            Button b = U<Button>("ExportBtn");
            b.Click += delegate { ContextMenu m = b.ContextMenu; m.PlacementTarget = b; m.Placement = PlacementMode.Bottom; m.IsOpen = true; };
            ExpItem(0).Click += delegate { ExportClientsPdf(); };
            ExpItem(1).Click += delegate { ExportClientsCsv(); };
        }
        MenuItem ExpItem(int i) { return (MenuItem)U<Button>("ExportBtn").ContextMenu.Items[i]; }

        string TramiteTitle() { return G.T.TitleText(); }

        string GetExportFile(string ext, string filter)
        {
            if (ExportFileHook != null) return ExportFileHook(ext, filter);
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = Docs.SafeName(string.Format("{0} {1} {2:yyyy-MM-dd}", L("navClients"), TramiteTitle(), App.Now()));
            dlg.DefaultExt = ext;
            string xd = ""; try { xd = Path.GetDirectoryName(Ps.Str(Config.Cur.ExcelPath)) ?? ""; } catch { }
            if (xd.Length > 0 && Directory.Exists(xd)) dlg.InitialDirectory = xd;
            dlg.Filter = filter;
            if (dlg.ShowDialog(Win) == true) return dlg.FileName;
            return null;
        }

        List<ClientRow> GridRows()
        {
            List<ClientRow> l = new List<ClientRow>();
            IEnumerable src = U<DataGrid>("ClientsGrid").ItemsSource;
            if (src != null) foreach (object o in src) { ClientRow r = o as ClientRow; if (r != null) l.Add(r); }
            return l;
        }

        static string CsvQ(object v) { string s = Ps.Str(v); return PsH.Is(s, "[;\"\r\n]") ? "\"" + s.Replace("\"", "\"\"") + "\"" : s; }

        public void ExportClientsCsv()
        {
            List<ClientRow> rows = GridRows();
            if (rows.Count == 0) { ShowInfo(L("exEmpty")); return; }
            string file = GetExportFile(".csv", "CSV (Excel)|*.csv");
            if (string.IsNullOrEmpty(file)) return;
            try
            {
                List<string> head = new List<string> { L("colId") };
                foreach (Field f in G.T.Fields) head.Add(f.LabelText());
                head.Add(L("colPend"));
                List<string> lines = new List<string>();
                List<string> hq = new List<string>(); foreach (string h in head) hq.Add(CsvQ(h));
                lines.Add(string.Join(";", hq.ToArray()));
                foreach (ClientRow r in rows)
                {
                    List<string> cells = new List<string> { CsvQ(r.Id) };
                    foreach (Field f in G.T.Fields)
                    {
                        object v = r.Data == null ? null : Store.Get(r.Data, f.Key);
                        if (v is DateTime) v = ((DateTime)v).ToString("dd/MM/yyyy");
                        else if (v is double) v = ((double)v).ToString("0.##", Ps.ES);
                        cells.Add(CsvQ(v));
                    }
                    cells.Add(CsvQ(r.Pending.ToString("0.00", Ps.ES)));
                    lines.Add(string.Join(";", cells.ToArray()));
                }
                App.Guard(file, "CSV");
                File.WriteAllLines(file, lines.ToArray(), new UTF8Encoding(true));
                App.Log(string.Format("exportó la lista a CSV ({0} clientes)", rows.Count));
                SetStatus(string.Format(L("exDone"), Path.GetFileName(file)));
                ShellOpen("explorer.exe", "/select,\"" + file + "\"");
            }
            catch (Exception e) { ShowError(L("errGeneric") + Environment.NewLine + e.Message); }
        }

        static readonly object[][] ListCols = {
            new object[] { "colId", 27.0 }, new object[] { "colName", 60.0 }, new object[] { "colDoc", 27.0 }, new object[] { "colTipo", 33.0 }, new object[] { "colStatus", 40.0 },
            new object[] { "colPay", 22.0 }, new object[] { "colPend", 24.0 }, new object[] { "colDays", 12.0 }, new object[] { "colResp", 28.0 } };

        internal static bool DrawClientList(Graphics g, int page)
        {
            Dictionary<string, object> d = Pdf.Data;
            double LE = 12, RE = 285;
            // cabecera
            Pdf.Rect(g, 0, 0, 297, 3, "#9E1820");
            bool hasLogo = Pdf.Logo(g, LE, 9, 13);
            double tx = hasLogo ? LE + 16 : LE;
            Pdf.Text(g, d["Company"], tx, 9.5, 150, 12, "Semibold", "#1E1A1B");
            Pdf.Text(g, d["Title"], tx, 16, 170, 9, "Regular", "#6E6466");
            Pdf.Text(g, d["Sub"], 150, 10.5, RE - 150, 8, "Regular", "#8A7F81", "Far");
            Pdf.Text(g, d["Sub2"], 150, 15.5, RE - 150, 8, "Regular", "#8A7F81", "Far");
            // cabecera de tabla
            double y = 27;
            Pdf.Rect(g, LE, y, RE - LE, 7, "#F3EFEF", 1.5);
            double x = LE + 2;
            foreach (object[] c in ListCols)
            {
                string k = (string)c[0]; double w = (double)c[1];
                Pdf.Text(g, Tx.L(k), x, y + 1.6, w - 3, 7.5, "Semibold", "#6E6466", k == "colPend" || k == "colDays" ? "Far" : "Near");
                x += w;
            }
            y += 8.5;
            List<ClientRow> rows = (List<ClientRow>)d["Rows"]; int per = (int)d["PerPage"];
            int start = page * per, end = Math.Min(rows.Count, start + per);
            for (int i = start; i < end; i++)
            {
                ClientRow r = rows[i];
                x = LE + 2;
                string[] vals = { Ps.Str(r.Id), Ps.Str(r.Name), Ps.Str(r.Doc), Ps.Str(r.Tipo), Ps.Str(r.Status), Ps.Str(r.Pay), Ps.Str(r.PendingText), Ps.Str(r.Days), Ps.Str(r.Resp) };
                for (int c = 0; c < ListCols.Length; c++)
                {
                    double w = (double)ListCols[c][1];
                    string col = Ps.Str(r.Color);
                    if (c == 4 && col.Length > 0 && col != "#FFFFFF")
                    {
                        Pdf.Rect(g, x, y + 1.6, 2.6, 2.6, col, 1.3);
                        Pdf.Text(g, vals[c], x + 4, y + 0.8, w - 7, 8, "Regular", "#2B2526");
                    }
                    else
                    {
                        string al = c == 6 || c == 7 ? "Far" : "Near";
                        string fc = c == 6 ? "#B42318" : "#2B2526";
                        string wt = c == 1 || c == 6 ? "Semibold" : "Regular";
                        Pdf.Text(g, vals[c], x, y + 0.8, w - 3, 8, wt, fc, al);
                    }
                    x += w;
                }
                Pdf.Line(g, LE, y + 6.2, RE, y + 6.2, "#F1EDED", 0.25);
                y += 6.4;
            }
            bool last = end >= rows.Count;
            if (last)
            {
                y += 3;
                Pdf.Rect(g, LE, y, RE - LE, 9, "#F7F5F5", 1.5);
                Pdf.Text(g, d["Total"], LE + 3, y + 2.2, 200, 9, "Semibold", "#1E1A1B");
                Pdf.Text(g, d["TotalPend"], 180, y + 2.2, RE - 183, 9, "Semibold", "#B42318", "Far");
            }
            Pdf.Text(g, string.Format(Tx.L("exPage"), page + 1, d["Pages"]), LE, 201, 100, 7.5, "Regular", "#A69C9E");
            Pdf.Text(g, "GESTEM MANAGER", 180, 201, RE - 180, 7.5, "Regular", "#A69C9E", "Far");
            return !last;
        }

        public void ExportClientsPdf()
        {
            List<ClientRow> rows = GridRows();
            if (rows.Count == 0) { ShowInfo(L("exEmpty")); return; }
            if (!Pdf.HasPrinter()) { ShowError(L("noPdfPrinter")); return; }
            string file = GetExportFile(".pdf", "PDF|*.pdf");
            if (string.IsNullOrEmpty(file)) return;
            Win.Cursor = Cursors.Wait;
            try
            {
                int per = 25;
                double sumPend = 0; int nPend = 0;
                foreach (ClientRow r in rows) if (r.Pending > 0) { sumPend += r.Pending; nPend++; }
                List<string> filt = new List<string>();
                string q = Ps.Str(U<TextBox>("SearchBox").Text).Trim(); if (q.Length > 0) filt.Add("«" + q + "»");
                ComboBox fb = U<ComboBox>("FilterBox"); if (fb.SelectedIndex > 0) filt.Add(Ps.Str(fb.SelectedItem));
                if (U<ToggleButton>("OnlyDebt").IsChecked == true) filt.Add(L("onlyDebt"));
                if (U<ToggleButton>("OnlyMine").IsChecked == true) filt.Add(L("onlyMine"));
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["Company"] = Ps.Str(Config.Cur.CompanyName).Trim().Length > 0 ? Ps.Str(Config.Cur.CompanyName) : "GESTEM MANAGER";
                d["Title"] = string.Format(L("exTitle"), TramiteTitle());
                d["Sub"] = App.Now().ToString("dd/MM/yyyy HH:mm") + (App.User.Length > 0 ? "  ·  " + App.User : "");
                d["Sub2"] = filt.Count > 0 ? L("exFilters") + ": " + string.Join(", ", filt.ToArray()) : L("exNoFilter");
                d["Rows"] = rows; d["PerPage"] = per; d["Pages"] = (int)Math.Max(1.0, Math.Ceiling(rows.Count / (double)per));
                d["Total"] = string.Format(L("exTotal"), rows.Count, nPend); d["TotalPend"] = L("colPend") + ": " + Ps.Money(sumPend);
                Pdf.Data = d;
                Pdf.Print(file, true, DrawClientList);
                App.Log(string.Format("exportó la lista a PDF ({0} clientes)", rows.Count));
                SetStatus(string.Format(L("exDone"), Path.GetFileName(file)));
                if (File.Exists(file)) ShellOpen(file, null);
            }
            catch (Exception e)
            {
                App.Log("exportar PDF: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Win.Cursor = null; }
        }
    }
}

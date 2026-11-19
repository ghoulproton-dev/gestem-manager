// Etapa 7 (sección 0): lo que faltaba antes de instalar GESTEM.exe, comparado con tools\pstest7.ps1 (mismas líneas en el resumen).
using System;
using System.Text;
using System.Text.RegularExpressions;

namespace Gestem
{
    public static class SelfTest7
    {
        // rutas de la prueba iguales en los dos: <T>\X en vez de <prueba>\cs o <prueba>\ps
        static string Norm(string s, string testDir) { return Regex.Replace(s.Replace(testDir, "<T>"), @"<T>\\(cs|ps)", @"<T>\X"); }

        public static void Run(MainWin w, StringBuilder o, string testDir)
        {
            MainWin.OpenHook = delegate(string s) { o.Append("OPEN " + Norm(s, testDir) + "\r\n"); };
            w.InfoHook = delegate(string m) { o.Append("INFO " + m + "\r\n"); };
            try
            {
                // abrir el Excel: sin archivo -> aviso; con archivo y sin Excel abierto con él -> se abre el archivo
                G.T = Tramite.ById("canje");
                string keepX = Config.Cur.ExcelPath;
                Config.Cur.ExcelPath = System.IO.Path.Combine(testDir, "no-existe.xlsx"); o.Append("XL0\r\n"); w.OpenExcelWindow(5);
                Config.Cur.ExcelPath = keepX; o.Append("XL1\r\n"); w.OpenExcelWindow(5); o.Append("XL2\r\n"); w.OpenExcelWindow(Hand.RowBase + 3);

                // factura PDF: se dibuja en imágenes (huella de los píxeles) en vez de imprimir; datos, nº y aviso, iguales que PowerShell
                App.Now = delegate { return new DateTime(2026, 11, 20, 10, 0, 0); };
                Pdf.PrintHook = delegate(string path, bool land, Pdf.DrawPage draw)
                {
                    o.Append("PDF " + Norm(path, testDir) + " land=" + land + " " + Pdf.Render(land, draw) + "\r\n");
                    o.Append("PDFD " + Dump(Pdf.Data) + "\r\n");
                };
                Config c = Config.Cur;
                c.InvName = "GESTORÍA PRUEBA"; c.InvNif = ""; c.InvActivity = "Gestión de trámites"; c.InvAddress = "Calle Uno, 1"; c.InvTown = "30700 Torre Pacheco";
                c.InvIban = "ES00 0000 0000 0000"; c.CompanyName = ""; c.CompanyLogo = ""; c.TasasCanje = "B=33; C=100; D=100";
                G.T = Tramite.ById("canje"); Data.LoadData(false, false);
                ClientRow withT = null, paidAll = null;
                foreach (ClientRow r in G.Rows)
                {
                    if (withT == null && r.PriceV > 0 && r.Pending > 0 && Ps.TasasFor(G.T, r.Data) > 0) withT = r;
                    if (paidAll == null && r.PriceV > 0 && r.Pending <= 0) paidAll = r;
                }
                o.Append("IN rows=" + G.Rows.Count + " conTasas=" + (withT == null ? "-" : withT.Id) + " pagado=" + (paidAll == null ? "-" : paidAll.Id) + "\r\n");
                o.Append("IN0 sin emisor\r\n"); w.NewClientInvoice(withT);
                c.InvNif = "X1234567L";
                o.Append("IN1 " + withT.Id + "\r\n"); w.NewClientInvoice(withT); o.Append("ST " + w.U<System.Windows.Controls.TextBlock>("StatusText").Text + "\r\n");
                o.Append("IN2 otra vez " + withT.Id + "\r\n"); w.NewClientInvoice(withT);
                o.Append("IN3 " + paidAll.Id + "\r\n"); w.NewClientInvoice(paidAll);
                double keepP = withT.PriceV; withT.PriceV = 0; o.Append("IN4 sin precio " + withT.Id + "\r\n"); w.NewClientInvoice(withT); withT.PriceV = keepP;
                G.T = Tramite.ById("traduccion"); Data.LoadData(false, false);
                ClientRow aut = null; foreach (ClientRow r in G.Rows) if (aut == null && r.PriceV > 0) aut = r;
                c.InvIban = ""; c.CompanyName = "EMPRESA X";
                o.Append("IN5 " + (aut == null ? "(ninguno)" : aut.Id) + "\r\n"); if (aut != null) w.NewClientInvoice(aut);
                // exportar la lista: vacía -> aviso; CSV (huella del archivo); PDF apaisado (huella de cada página); diálogo cancelado -> nada
                G.T = Tramite.ById("canje"); Data.LoadData(false, false); c.CompanyName = ""; string keepU = App.User; App.User = "PRUEBA";
                System.Windows.Controls.DataGrid eg = w.U<System.Windows.Controls.DataGrid>("ClientsGrid");
                string expPath = null; MainWin.ExportFileHook = delegate(string ext, string filter) { o.Append("EXF " + ext + " " + filter + "\r\n"); return expPath; };
                try
                {
                    eg.ItemsSource = new System.Collections.Generic.List<ClientRow>(); o.Append("EX0 vacía\r\n"); w.ExportClientsCsv(); w.ExportClientsPdf();
                    eg.ItemsSource = G.Rows;
                    expPath = System.IO.Path.Combine(testDir, @"cs\lista.csv"); o.Append("EX1 csv\r\n"); w.ExportClientsCsv();
                    byte[] cb = System.IO.File.ReadAllBytes(expPath);
                    using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create()) o.Append("CSV bytes=" + cb.Length + " md5=" + BitConverter.ToString(md5.ComputeHash(cb)).Replace("-", "").Substring(0, 12) + "\r\n");
                    o.Append("ST " + w.U<System.Windows.Controls.TextBlock>("StatusText").Text + "\r\n");
                    expPath = System.IO.Path.Combine(testDir, @"cs\lista.pdf"); o.Append("EX2 pdf\r\n"); w.ExportClientsPdf();
                    expPath = null; o.Append("EX3 cancelado\r\n"); w.ExportClientsCsv(); w.ExportClientsPdf();
                }
                finally { MainWin.ExportFileHook = null; App.User = keepU; }
            }
            finally { MainWin.OpenHook = null; w.InfoHook = null; Pdf.PrintHook = null; App.Now = delegate { return DateTime.Now; }; G.T = Tramite.ById("canje"); Data.LoadData(false, false); }
        }

        // los datos de la factura o la lista, en orden de clave (como pstest7)
        static string Dump(System.Collections.Generic.Dictionary<string, object> d)
        {
            System.Collections.Generic.List<string> keys = new System.Collections.Generic.List<string>(d.Keys); keys.Sort(StringComparer.Ordinal);
            System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
            foreach (string k in keys)
            {
                object v = d[k]; string s;
                if (k == "Rows") s = "n=" + ((System.Collections.ICollection)v).Count;
                else if (v == null) s = "";
                else if (v is double) s = ((double)v).ToString("0.##", Ps.Inv);
                else if (v is string[][]) { System.Collections.Generic.List<string> l = new System.Collections.Generic.List<string>(); foreach (string[] a in (string[][])v) l.Add(a[0] + "=" + a[1]); s = string.Join(";", l.ToArray()); }
                else s = Ps.Str(v);
                parts.Add(k + "=" + s);
            }
            return string.Join(" | ", parts.ToArray());
        }
    }
}

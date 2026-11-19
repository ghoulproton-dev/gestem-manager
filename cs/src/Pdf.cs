// PDF con la impresora «Microsoft Print to PDF» de Windows, sin diálogos. Copiado de Test-PdfPrinter, New-PdfFont, PdfText, PdfWrap,
// PdfRect, PdfLine, PdfLogo, PdfTextW, PdfLV e Invoke-PdfPrint de GESTEM.ps1 (mismas medidas en mm, mismos colores y fuentes).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Threading;

namespace Gestem
{
    public static class Pdf
    {
        public const string Printer = "Microsoft Print to PDF";
        public delegate bool DrawPage(Graphics g, int page);
        public static Dictionary<string, object> Data = new Dictionary<string, object>();    // $script:PdfData
        // pruebas: en vez de imprimir, se dibuja en imágenes y se compara con PowerShell (Invoke-PdfPrint de pstest7)
        public static Action<string, bool, DrawPage> PrintHook;

        public static bool HasPrinter()
        {
            try { foreach (string p in PrinterSettings.InstalledPrinters) if (string.Equals(p, Printer, StringComparison.OrdinalIgnoreCase)) return true; } catch { }
            return false;
        }

        static Font NewFont(double size, string weight)
        {
            string fam = "Segoe UI"; FontStyle st = FontStyle.Regular;
            if (weight == "Semibold") fam = "Segoe UI Semibold";
            else if (weight == "Bold") st = FontStyle.Bold;
            else if (weight == "Light") fam = "Segoe UI Light";
            return new Font(fam, (float)size, st, GraphicsUnit.Point);
        }
        static Color C(string hex) { return ColorTranslator.FromHtml(hex); }

        public static void Text(Graphics g, object text, double x, double y, double w, double size, string weight = "Regular", string color = "#1E1A1B", string align = "Near")
        {
            using (Font font = NewFont(size, weight))
            using (SolidBrush br = new SolidBrush(C(color)))
            using (StringFormat sf = new StringFormat())
            {
                sf.Alignment = (StringAlignment)Enum.Parse(typeof(StringAlignment), align, true);
                sf.Trimming = StringTrimming.EllipsisCharacter;
                sf.FormatFlags = StringFormatFlags.NoWrap;
                double h = size * 0.55 + 1.5;
                g.DrawString(Ps.Str(text), font, br, new RectangleF((float)x, (float)y, (float)w, (float)h), sf);
            }
        }
        public static void Wrap(Graphics g, object text, double x, double y, double w, double h, double size, string color = "#6E6466")
        {
            using (Font font = NewFont(size, "Regular"))
            using (SolidBrush br = new SolidBrush(C(color)))
                g.DrawString(Ps.Str(text), font, br, new RectangleF((float)x, (float)y, (float)w, (float)h));
        }
        public static void Rect(Graphics g, double x, double y, double w, double h, string fill, double radius = 0)
        {
            using (SolidBrush br = new SolidBrush(C(fill)))
            {
                if (radius <= 0) { g.FillRectangle(br, (float)x, (float)y, (float)w, (float)h); return; }
                float d = (float)(radius * 2);
                using (GraphicsPath gp = new GraphicsPath())
                {
                    gp.AddArc((float)x, (float)y, d, d, 180, 90);
                    gp.AddArc((float)(x + w - d), (float)y, d, d, 270, 90);
                    gp.AddArc((float)(x + w - d), (float)(y + h - d), d, d, 0, 90);
                    gp.AddArc((float)x, (float)(y + h - d), d, d, 90, 90);
                    gp.CloseFigure();
                    g.FillPath(br, gp);
                }
            }
        }
        public static void Line(Graphics g, double x1, double y1, double x2, double y2, string color = "#E6E1E2", double width = 0.3)
        {
            using (Pen pen = new Pen(C(color), (float)width)) g.DrawLine(pen, (float)x1, (float)y1, (float)x2, (float)y2);
        }
        public static bool Logo(Graphics g, double x, double y, double size)
        {
            string p = Ps.Str(Config.Cur.CompanyLogo);
            if (p.Length == 0 || !File.Exists(p)) return false;
            try { using (Image img = Image.FromFile(p)) g.DrawImage(img, new RectangleF((float)x, (float)y, (float)size, (float)size)); return true; }
            catch { return false; }
        }
        public static double TextW(Graphics g, object text, double size, string weight = "Regular")
        {
            using (Font font = NewFont(size, weight)) return (double)g.MeasureString(Ps.Str(text), font).Width;
        }
        // «Etiqueta: valor» con la etiqueta en negrita
        public static void LV(Graphics g, string label, string value, double x, double y, double w, double size, string lc = "#0F172A", string vc = "#334155")
        {
            Text(g, label, x, y, w, size, "Semibold", lc);
            double lw = TextW(g, label, size, "Semibold") + 0.6;
            Text(g, value, x + lw, y, Math.Max(5, w - lw), size, "Regular", vc);
        }

        // Invoke-PdfPrint: draw recibe (g, nº de página) y devuelve true si hay más páginas; los datos van en Pdf.Data
        public static void Print(string path, bool landscape, DrawPage draw)
        {
            if (PrintHook != null) { PrintHook(path, landscape, draw); return; }
            App.Guard(path, "PDF");
            string err = null; int page = 0;
            using (PrintDocument doc = new PrintDocument())
            {
                doc.PrinterSettings.PrinterName = Printer;
                doc.PrinterSettings.PrintToFile = true;
                doc.PrinterSettings.PrintFileName = path;
                doc.DocumentName = Path.GetFileNameWithoutExtension(path);
                doc.PrintController = new StandardPrintController();
                foreach (PaperSize ps in doc.PrinterSettings.PaperSizes) if (ps.Kind == PaperKind.A4) { doc.DefaultPageSettings.PaperSize = ps; break; }
                doc.DefaultPageSettings.Landscape = landscape;
                doc.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                doc.PrintPage += delegate(object s, PrintPageEventArgs e)
                {
                    try
                    {
                        Graphics g = e.Graphics;
                        g.PageUnit = GraphicsUnit.Millimeter;
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.TranslateTransform((float)(-e.PageSettings.HardMarginX * 0.254), (float)(-e.PageSettings.HardMarginY * 0.254));
                        bool more = draw(g, page);
                        page++;
                        e.HasMorePages = more;
                    }
                    catch (Exception x) { err = x.Message; e.HasMorePages = false; }
                };
                if (File.Exists(path)) File.Delete(path);
                doc.Print();
            }
            if (err != null) throw new GestemException(err);
            // la impresora escribe el archivo en segundo plano: esperar a que exista y a que el tamaño deje de cambiar
            for (int i = 0; i < 60 && !File.Exists(path); i++) Thread.Sleep(150);
            if (!File.Exists(path)) throw new GestemException(Tx.L("pdfFail"));
            long last = -1;
            for (int i = 0; i < 40; i++) { long len = new FileInfo(path).Length; if (len > 0 && len == last) break; last = len; Thread.Sleep(200); }
        }

        // pruebas: dibuja las páginas en imágenes A4 (100 ppp) y devuelve una huella de los píxeles de cada una
        public static string Render(bool landscape, DrawPage draw)
        {
            int wpx = landscape ? 1169 : 827, hpx = landscape ? 827 : 1169;
            List<string> hs = new List<string>(); int page = 0; bool more;
            do
            {
                using (Bitmap bmp = new Bitmap(wpx, hpx, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    bmp.SetResolution(100, 100);
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.White);
                        g.PageUnit = GraphicsUnit.Millimeter; g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit; g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        more = draw(g, page);
                    }
                    System.Drawing.Imaging.BitmapData bd = bmp.LockBits(new Rectangle(0, 0, wpx, hpx), System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    byte[] px = new byte[bd.Stride * hpx]; System.Runtime.InteropServices.Marshal.Copy(bd.Scan0, px, 0, px.Length); bmp.UnlockBits(bd);
                    using (System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create()) hs.Add(BitConverter.ToString(md5.ComputeHash(px)).Replace("-", "").Substring(0, 12));
                }
                page++;
            } while (more && page < 40);
            return "pages=" + page + " px=" + string.Join(",", hs.ToArray());
        }
    }
}

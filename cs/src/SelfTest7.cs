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
            try
            {
                // abrir el Excel: sin archivo -> aviso; con archivo y sin Excel abierto con él -> se abre el archivo
                G.T = Tramite.ById("canje");
                string keepX = Config.Cur.ExcelPath;
                Config.Cur.ExcelPath = System.IO.Path.Combine(testDir, "no-existe.xlsx"); o.Append("XL0\r\n"); w.OpenExcelWindow(5);
                Config.Cur.ExcelPath = keepX; o.Append("XL1\r\n"); w.OpenExcelWindow(5); o.Append("XL2\r\n"); w.OpenExcelWindow(Hand.RowBase + 3);
            }
            finally { MainWin.OpenHook = null; }
        }
    }
}

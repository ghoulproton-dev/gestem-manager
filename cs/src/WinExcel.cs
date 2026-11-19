// «Abrir el Excel» (botón, clic derecho y «abrir el Excel después de guardar»). Copiado de Open-ExcelWindow de GESTEM.ps1:
// si el Excel ya está abierto, se trae al frente en la pestaña del trámite y la fila del cliente; si no, se abre el archivo.
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace Gestem
{
    public partial class MainWin
    {
        // pruebas: en vez de abrir un archivo o el explorador, se apunta aquí (Start-Process de pstest)
        public static Action<string> OpenHook;

        // Start-Process: abre el archivo con su programa (o el explorador con argumentos)
        public static void ShellOpen(string file, string args)
        {
            if (OpenHook != null) { OpenHook(args == null ? file : file + " " + args); return; }
            ProcessStartInfo p = new ProcessStartInfo(file); if (args != null) p.Arguments = args; p.UseShellExecute = true;
            Process.Start(p);
        }

        void WireExcel()
        {
            U<Button>("OpenExcelBtn").Click += delegate
            {
                ClientRow sel = U<DataGrid>("ClientsGrid").SelectedItem as ClientRow;
                OpenExcelWindow(sel != null ? sel.Row : 0);
            };
            cmExcel.Click += delegate
            {
                ClientRow r = SelectedClient(); if (r == null) return;
                OpenExcelWindow(r.Row < Hand.RowBase ? r.Row : 0);
            };
        }

        public void OpenExcelWindow(int row)
        {
            string path = Config.Cur.ExcelPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) { ShowError(L("errNoFile")); return; }
            try
            {
                object xl = null;
                try { xl = Marshal.GetActiveObject("Excel.Application"); } catch { }
                if (xl != null)
                {
                    string full = Path.GetFullPath(path);
                    foreach (object w in Com.Items(Com.Get(xl, "Workbooks")))
                    {
                        if (!string.Equals(Ps.Str(Com.Get(w, "FullName")), full, StringComparison.OrdinalIgnoreCase)) continue;
                        Com.Set(xl, "Visible", true); Com.Call(w, "Activate");
                        object tab = ExcelCom.FindTramiteSheet(w, G.T);
                        if (row >= Hand.RowBase && Hand.HandTabs[G.T.Id] != null)
                        {
                            tab = ExcelCom.FindSheet(w, (string)Hand.HandTabs[G.T.Id]); row -= Hand.RowBase;
                            if (tab != null) { Com.Call(tab, "Activate"); Com.Call(xl, "Goto", Com.Range(tab, "A" + row)); }
                            return;
                        }
                        if (row > 0 && tab != null) { Com.Call(tab, "Activate"); Com.Call(xl, "Goto", Com.Range(tab, "B" + row)); }
                        return;
                    }
                }
            }
            catch { }
            try { ShellOpen(path, null); } catch (Exception e) { ShowError(L("errGeneric") + Environment.NewLine + e.Message); }
        }
    }
}

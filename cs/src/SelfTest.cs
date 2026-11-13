// Prueba de igualdad con la versión PowerShell: lee los datos de una carpeta de prueba, calcula cada fila
// y vuelve a guardar los archivos. tools\pstest.ps1 hace lo mismo con las funciones de GESTEM.ps1.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Gestem
{
    public static class SelfTest
    {
        public static string N(double v) { return v.ToString("0.##", Ps.Inv); }

        public static int Run(string testDir)
        {
            StringBuilder o = new StringBuilder();
            App.DataDir = testDir; App.BackupDir = Path.Combine(testDir, "copias"); App.LogPath = Path.Combine(testDir, "cs_log.txt");
            Directory.CreateDirectory(App.BackupDir);
            Config.Cur = new Config();
            Config.Cur.ExcelPath = Path.Combine(testDir, "cs", "x.xlsx");
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = Store.Read(t);
                if (s == null) { o.Append("T " + t.Id + " (sin datos)\r\n"); continue; }
                TramiteData d = Store.BuildData(t, s);
                double pend = 0, price = 0, paid = 0, net = 0;
                foreach (ClientRow r in d.Rows) { pend += r.Pending; price += r.PriceV; paid += r.PaidV; net += r.NetV; }
                o.Append("T " + t.Id + " rows=" + d.Rows.Count + " firstEmpty=" + (d.FirstEmpty.HasValue ? d.FirstEmpty.Value.ToString() : "") + " nextId=" + d.NextId
                    + " pend=" + N(pend) + " price=" + N(price) + " paid=" + N(paid) + " net=" + N(net) + "\r\n");
                foreach (ClientRow r in d.Rows)
                    o.Append("R " + t.Id + " " + r.Row + " " + r.Id + "|" + r.Name + "|" + r.Doc + "|" + r.Tipo + "|" + r.Status + "|" + r.Color + "|" + r.Pay + "|"
                        + N(r.Pending) + "|" + r.PendingText + "|" + Ps.Str(r.Days) + "|" + r.Loc + "|" + r.Resp + "|" + N(r.PriceV) + "|" + N(r.PaidV) + "|" + N(r.NetV) + "|"
                        + (r.Fecha.HasValue ? r.Fecha.Value.ToString("yyyy-MM-dd", Ps.Inv) : "") + "\r\n");
                List<string> keys = new List<string>(d.Lists.Keys); keys.Sort(StringComparer.Ordinal);
                foreach (string k in keys) o.Append("L " + t.Id + " " + k + "=" + string.Join("|", d.Lists[k]) + "\r\n");
                // resolver un registro como al guardar
                foreach (ClientRec c in s.Clients)
                {
                    Dictionary<string, object> fv = new Dictionary<string, object>(c.F);
                    Store.Resolve(t, fv, s.Meta);
                    o.Append("V " + t.Id + " " + c.Row + " price=" + Ps.Str(fv.ContainsKey("price") ? fv["price"] : null) + " paid=" + Ps.Str(fv.ContainsKey("paid") ? fv["paid"] : null) + "\r\n");
                }
                Store.Write(t, s);
            }
            File.WriteAllText(Path.Combine(testDir, "cs_summary.txt"), o.ToString(), new UTF8Encoding(false));
            return 0;
        }
    }
}

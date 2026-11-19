// Libro de facturas («GESTEM datos\facturas.json»): Get-InvoiceBook, Save-InvoiceBook y Set-InvoiceNext de GESTEM.ps1.
// shortcut: Get-InvoiceNumber (sacar una factura) llega con las facturas PDF en la etapa 5.
using System;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Gestem
{
    public class InvoiceBook { public int Next = 5; public Hashtable Issued = PsH.New(); }

    public static class Invoices
    {
        public static InvoiceBook Get()
        {
            InvoiceBook b = new InvoiceBook();
            string d = Store.Dir();
            if (d == null) return b;
            string p = Path.Combine(d, "facturas.json");
            if (File.Exists(p))
            {
                try
                {
                    object j = Json.Parse(File.ReadAllText(p, Encoding.UTF8));
                    object n = Json.Get(j, "next"); if (n != null && Convert.ToDouble(n) != 0) b.Next = Convert.ToInt32(n);
                    Dictionary<string, object> iss = Json.Obj(Json.Get(j, "issued"));
                    if (iss != null) foreach (KeyValuePair<string, object> kv in iss) b.Issued[kv.Key] = Ps.Str(kv.Value);
                }
                catch (Exception e) { App.Log("facturas.json ilegible: " + e.Message); }
            }
            return b;
        }

        public static void Save(InvoiceBook b)
        {
            string d = Store.Dir();
            if (d == null) throw new GestemException(Tx.L("errNoFile"));
            string p = Path.Combine(d, "facturas.json");
            App.Guard(p, "facturas");
            Directory.CreateDirectory(d);
            OrderedDictionary o = new OrderedDictionary();
            o["app"] = "GESTEM MANAGER"; o["next"] = b.Next; o["issued"] = b.Issued;
            File.WriteAllText(p, Json.Write(o), new UTF8Encoding(true));
        }

        public static void SetNext(int n)
        {
            string lk = Store.EnterLock();
            try
            {
                InvoiceBook b = Get();
                if (b.Next != n) { b.Next = n; Save(b); App.Log("próximo nº de factura: " + n.ToString("000")); }
            }
            finally { Store.ExitLock(lk); }
        }
    }
}

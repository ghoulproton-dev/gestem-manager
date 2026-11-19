// Factura PDF del cliente (botón «Recibo» y clic derecho). Copiado de New-ClientInvoice, Get-InvoiceConcept y Draw-Invoice de
// GESTEM.ps1: mismo diseño, nº común a todos los PC (facturas.json), tasas DGT del canje como suplido sin IVA.
// (New-ClientReceipt / Draw-Receipt no se trasladan: en GESTEM.ps1 nada los llama.)
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Controls;
using System.Windows.Input;

namespace Gestem
{
    public partial class MainWin
    {
        void WireInvoice()
        {
            U<Button>("ReceiptBtn").Click += delegate { ClientRow r = SelectedClient(); if (r != null) NewClientInvoice(r); };
            U<MenuItem>("CmReceipt").Click += delegate { ClientRow r = SelectedClient(); if (r != null) NewClientInvoice(r); };
        }

        static string[] Pair(string a, string b) { return new string[] { a, b }; }

        // Get-InvoiceConcept
        static void InvoiceConcept(Tramite t, ClientRow r, out string title, out string desc)
        {
            string tipo = Ps.Str(r.Tipo).Trim();
            if (t.Id == "canje")
            {
                desc = "Asesoramiento, preparación de documentación y solicitud telemática/presencial ante la DGT.";
                if (tipo.Length > 0) desc += " Permiso: " + tipo + ".";
                title = "Gestión y tramitación de canje de permiso de conducir";
                return;
            }
            string extra = tipo.Length > 0 ? " (" + tipo + ")" : "";
            title = string.Format("Gestión y tramitación: {0}{1}", t.TitleEs, extra);
            desc = "Asesoramiento, preparación de documentación y presentación del trámite.";
        }

        public void NewClientInvoice(ClientRow r)
        {
            if (r == null) return;
            Config cfg = Config.Cur;
            if (Ps.Str(cfg.InvName).Trim().Length == 0 || Ps.Str(cfg.InvNif).Trim().Length == 0) { ShowInfo(L("invNoIssuer")); return; }
            double total = r.PriceV;
            if (total <= 0) { ShowInfo(L("invNoPrice")); return; }
            if (!Pdf.HasPrinter()) { ShowError(L("noPdfPrinter")); return; }
            Win.Cursor = Cursors.Wait;
            try
            {
                Tramite t = G.T;
                string id = Ps.Str(r.Id).Trim(); if (id.Length == 0) id = "F" + r.Row;
                DateTime date; bool isNew;
                string no = Invoices.Number(t.Id + "|" + id, out date, out isNew);
                double bas, iva, tx; Ps.SplitIva(total, Ps.TasasFor(t, r.Data), out bas, out iva, out tx);
                double pend = r.Pending;
                string forma = Ps.Str(Store.Get(r.Data, "forma")).Trim();
                bool cash = PsH.Is(forma, "(?i)efectivo|cash|esp[eè]ces");
                List<string[]> pay = new List<string[]>();
                pay.Add(Pair("Forma de pago", forma.Length > 0 ? forma : "Transferencia bancaria / Bizum"));
                if (Ps.Str(cfg.InvIban).Trim().Length > 0 && !cash) pay.Add(Pair("IBAN", Ps.Str(cfg.InvIban).Trim()));
                pay.Add(Pair("Concepto", string.Format("Factura {0} - {1}", no, Ps.Str(r.Name))));
                List<string[]> issuer = new List<string[]> { Pair("NIF/NIE", Ps.Str(cfg.InvNif)) };
                if (Ps.Str(cfg.InvAddress).Trim().Length > 0) issuer.Add(Pair("Domicilio", Ps.Str(cfg.InvAddress)));
                if (Ps.Str(cfg.InvTown).Trim().Length > 0) issuer.Add(Pair("Población", Ps.Str(cfg.InvTown)));
                List<string[]> client = new List<string[]> { Pair("NIF/NIE", Ps.Str(r.Doc).Trim().Length > 0 ? Ps.Str(r.Doc) : "—") };
                string nat = Ps.Str(Store.Get(r.Data, "nat")).Trim(); if (nat.Length > 0) client.Add(Pair("Nacionalidad", nat));
                string tel = Ps.Str(Store.Get(r.Data, "tel")).Trim(); if (tel.Length > 0) client.Add(Pair("Teléfono", tel));
                client.Add(Pair("Expediente", id));
                string ct, cd; InvoiceConcept(t, r, out ct, out cd);
                Dictionary<string, object> d = new Dictionary<string, object>();
                d["No"] = no; d["Date"] = date.ToString("dd/MM/yyyy");
                d["Issuer"] = Ps.Str(cfg.InvName); d["Activity"] = Ps.Str(cfg.InvActivity); d["IssuerLines"] = issuer.ToArray();
                d["Client"] = Ps.Str(r.Name); d["ClientLines"] = client.ToArray();
                d["ConceptTitle"] = ct; d["ConceptDesc"] = cd;
                d["Base"] = bas; d["Iva"] = iva; d["Tasas"] = tx; d["Total"] = total; d["Pend"] = pend; d["PaidAmt"] = Math.Max(0.0, total - pend);
                d["Company"] = Ps.Str(cfg.CompanyName).Trim().Length > 0 ? Ps.Str(cfg.CompanyName) : Ps.Str(cfg.InvName);
                d["PayLines"] = pay.ToArray();
                Pdf.Data = d;
                string dir = Docs.ClientDir(t, r, true);
                string file = Path.Combine(dir, string.Format("Factura {0} - {1}.pdf", no, Docs.SafeName(r.Name)));
                Pdf.Print(file, false, DrawInvoice);
                App.Log(string.Format("creó la factura {0} ({1})", no, id));
                Docs.UpdateMap(); RefreshClientGrid();
                if (File.Exists(file)) ShellOpen(file, null);
                SetStatus(string.Format(L("invDone"), no));
            }
            catch (Exception e)
            {
                App.Log("factura: " + e.Message);
                ShowError(L("errGeneric") + Environment.NewLine + e.Message);
            }
            finally { Win.Cursor = null; }
        }

        // Draw-Invoice: mismo diseño que el recibo de GESTEM (banda roja, logo, cajas grises) con los datos de la factura
        internal static bool DrawInvoice(Graphics g, int page)
        {
            Dictionary<string, object> d = Pdf.Data;
            double LE = 18, RE = 192, W = RE - LE;
            string red = "#9E1820", dark = "#1E1A1B", grey = "#8A7F81", mid = "#4B4345";
            // cabecera
            Pdf.Rect(g, 0, 0, 210, 4, red);
            bool hasLogo = Pdf.Logo(g, LE, 14, 20);
            double tx = hasLogo ? LE + 25 : LE;
            Pdf.Text(g, d["Company"], tx, 16.5, 90, 15, "Semibold", dark);
            Pdf.Text(g, d["Activity"], tx, 25, 90, 9, "Regular", grey);
            Pdf.Text(g, "FACTURA", 110, 14, RE - 110, 24, "Light", red, "Far");
            Pdf.Text(g, string.Format("Nº {0}", d["No"]), 110, 27, RE - 110, 10, "Semibold", dark, "Far");
            Pdf.Text(g, string.Format("Fecha de emisión: {0}", d["Date"]), 110, 32.5, RE - 110, 9, "Regular", "#6E6466", "Far");
            Pdf.Line(g, LE, 44, RE, 44, "#E6E1E2", 0.35);
            // emisor y cliente
            double y = 52;
            Pdf.Text(g, "EMISOR", LE, y, 80, 7.5, "Semibold", red);
            Pdf.Text(g, "CLIENTE", 110, y, 80, 7.5, "Semibold", red);
            Pdf.Text(g, d["Issuer"], LE, y + 5, 88, 12, "Semibold", dark);
            Pdf.Text(g, d["Client"], 110, y + 5, 82, 12, "Semibold", dark);
            int i = 0;
            foreach (string[] ln in (string[][])d["IssuerLines"])
            {
                double yy = y + 13 + i * 6;
                Pdf.Text(g, ln[0], LE, yy, 30, 8.5, "Regular", grey);
                Pdf.Text(g, ln[1], LE + 28, yy, 62, 9, "Semibold", "#2B2526");
                i++;
            }
            i = 0;
            foreach (string[] ln in (string[][])d["ClientLines"])
            {
                double yy = y + 13 + i * 6;
                Pdf.Text(g, ln[0], 110, yy, 32, 8.5, "Regular", grey);
                Pdf.Text(g, ln[1], 140, yy, 52, 9, "Semibold", "#2B2526");
                i++;
            }
            // concepto
            double cy = 88;
            Pdf.Text(g, "CONCEPTO", LE, cy, 80, 7.5, "Semibold", red);
            Pdf.Rect(g, LE, cy + 5, W, 8, "#F3EFEF", 1.5);
            Pdf.Text(g, "Descripción", LE + 4, cy + 6.8, 90, 7.5, "Semibold", "#6E6466");
            Pdf.Text(g, "Cant.", LE + 100, cy + 6.8, 22, 7.5, "Semibold", "#6E6466", "Center");
            Pdf.Text(g, "Precio unid.", LE + 122, cy + 6.8, 26, 7.5, "Semibold", "#6E6466", "Far");
            Pdf.Text(g, "Importe", LE + 148, cy + 6.8, W - 152, 7.5, "Semibold", "#6E6466", "Far");
            double ry = cy + 17;
            Pdf.Text(g, d["ConceptTitle"], LE + 4, ry, 94, 10, "Semibold", dark);
            Pdf.Wrap(g, d["ConceptDesc"], LE + 4, ry + 5.5, 94, 11, 8, "#6E6466");
            Pdf.Text(g, "1", LE + 100, ry + 2.5, 22, 9.5, "Regular", mid, "Center");
            Pdf.Text(g, Ps.Money(d["Base"]), LE + 122, ry + 2.5, 26, 9.5, "Regular", mid, "Far");
            Pdf.Text(g, Ps.Money(d["Base"]), LE + 148, ry + 2.5, W - 152, 9.5, "Semibold", dark, "Far");
            // tasas DGT (canje): suplido pagado en nombre del cliente, sin IVA
            bool hasT = d["Tasas"] != null && (double)d["Tasas"] > 0;
            double ly2 = ry + 18;
            if (hasT)
            {
                Pdf.Line(g, LE, ry + 18, RE, ry + 18, "#EEE9EA", 0.3);
                double ry2 = ry + 22;
                Pdf.Text(g, "Tasas DGT (suplido)", LE + 4, ry2, 94, 10, "Semibold", dark);
                Pdf.Text(g, "Pagadas a la DGT en nombre del cliente · sin IVA", LE + 4, ry2 + 5.5, 94, 8, "Regular", "#6E6466");
                Pdf.Text(g, "1", LE + 100, ry2 + 2.5, 22, 9.5, "Regular", mid, "Center");
                Pdf.Text(g, Ps.Money(d["Tasas"]), LE + 122, ry2 + 2.5, 26, 9.5, "Regular", mid, "Far");
                Pdf.Text(g, Ps.Money(d["Tasas"]), LE + 148, ry2 + 2.5, W - 152, 9.5, "Semibold", dark, "Far");
                ly2 = ry2 + 13;
            }
            Pdf.Line(g, LE, ly2, RE, ly2, "#EEE9EA", 0.3);
            // importes
            double by = ly2 + 6;
            double bh = hasT ? 54 : 46;
            Pdf.Rect(g, LE, by, W, bh, "#F7F5F5", 3);
            Pdf.Text(g, "Base imponible", LE + 8, by + 8, 100, 10, "Regular", mid);
            Pdf.Text(g, Ps.Money(d["Base"]), 110, by + 8, RE - 118, 10, "Semibold", dark, "Far");
            Pdf.Text(g, "IVA (21 %)", LE + 8, by + 16, 100, 10, "Regular", mid);
            Pdf.Text(g, Ps.Money(d["Iva"]), 110, by + 16, RE - 118, 10, "Semibold", dark, "Far");
            double tl = 26;
            if (hasT)
            {
                Pdf.Text(g, "Suplidos: tasas DGT (sin IVA)", LE + 8, by + 24, 100, 10, "Regular", mid);
                Pdf.Text(g, Ps.Money(d["Tasas"]), 110, by + 24, RE - 118, 10, "Semibold", dark, "Far");
                tl = 34;
            }
            Pdf.Line(g, LE + 8, by + tl, RE - 8, by + tl, "#DDD6D7", 0.3);
            Pdf.Text(g, "TOTAL FACTURA", LE + 8, by + tl + 5, 100, 12, "Semibold", dark);
            Pdf.Text(g, Ps.Money(d["Total"]), 100, by + tl + 4, RE - 108, 16, "Semibold", red, "Far");
            // forma de pago
            double py = by + bh + 8;
            Pdf.Text(g, "PAGO", LE, py, 80, 7.5, "Semibold", red);
            i = 0;
            foreach (string[] ln in (string[][])d["PayLines"])
            {
                double yy = py + 5.5 + i * 6;
                Pdf.Text(g, ln[0], LE, yy, 40, 8.5, "Regular", grey);
                Pdf.Text(g, ln[1], LE + 40, yy, 90, 9, "Semibold", "#2B2526");
                i++;
            }
            // sello de pago
            double pend = (double)d["Pend"], paid = (double)d["PaidAmt"];
            string stamp = pend <= 0 ? "PAGADA" : paid > 0 ? "PAGO PARCIAL" : "";
            if (stamp.Length > 0)
            {
                bool ok = pend <= 0;
                Pdf.Rect(g, RE - 52, py + 4, 52, 10, ok ? "#E7F4EC" : "#FDF3E1", 5);
                Pdf.Text(g, stamp, RE - 52, py + 6.3, 52, 9, "Semibold", ok ? "#1F7A45" : "#B7791F", "Center");
            }
            // firma
            Pdf.Line(g, LE, 235, LE + 70, 235, "#B9B0B2", 0.3);
            Pdf.Text(g, "Firma y sello", LE, 237, 70, 8, "Regular", grey);
            // pie
            Pdf.Line(g, LE, 280, RE, 280, "#EEE9EA", 0.3);
            Pdf.Text(g, string.Format("Factura Nº {0} — {1}", d["No"], d["Issuer"]), LE, 282.5, 110, 7.5, "Regular", "#A69C9E");
            Pdf.Text(g, Tx.L("rcFooter"), 120, 282.5, RE - 120, 7.5, "Regular", "#A69C9E", "Far");
            return false;
        }
    }
}

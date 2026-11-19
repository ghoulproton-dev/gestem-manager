// Comportamientos de PowerShell que GESTEM usa en todas partes ("$v", importes, fechas…),
// copiados tal cual para que la versión C# dé exactamente los mismos resultados.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Gestem
{
    public static class Ps
    {
        public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        public static readonly CultureInfo ES = CultureInfo.GetCultureInfo("es-ES");

        // "$v" de PowerShell
        public static string Str(object v)
        {
            if (v == null) return "";
            if (v is string) return (string)v;
            if (v is bool) return (bool)v ? "True" : "False";
            if (v is double) return ((double)v).ToString(Inv);
            if (v is float) return ((float)v).ToString(Inv);
            if (v is decimal) return ((decimal)v).ToString(Inv);
            if (v is DateTime) return ((DateTime)v).ToString("G", Inv);
            if (v is IConvertible) return Convert.ToString(v, Inv);
            IEnumerable en = v as IEnumerable;
            if (en != null)
            {
                StringBuilder sb = new StringBuilder();
                foreach (object x in en) { if (sb.Length > 0) sb.Append(' '); sb.Append(Str(x)); }
                return sb.ToString();
            }
            return v.ToString();
        }

        public static bool Blank(object v) { return v == null || Str(v) == ""; }
        public static bool IsNum(object v) { return v is double || v is int || v is long || v is decimal || v is float || v is short; }

        // Parse-Money
        public static double? ParseMoney(string s)
        {
            s = (s ?? "").Replace("€", "").Replace(" ", "").Trim();
            if (s.Length == 0) return null;
            if (s.Contains(".") && s.Contains(",")) s = s.Replace(".", "").Replace(",", "."); else s = s.Replace(",", ".");
            double d;
            if (double.TryParse(s, NumberStyles.Float, Inv, out d)) return d;
            return null;
        }

        // ConvertTo-CellNumber
        public static double? Num(object v)
        {
            if (Blank(v)) return null;
            if (IsNum(v)) return Convert.ToDouble(v, Inv);
            return ParseMoney(Str(v));
        }

        // ConvertTo-CellDate (número de Excel)
        public static double? CellDate(object v)
        {
            if (Blank(v)) return null;
            if (v is DateTime) return ((DateTime)v).ToOADate();
            if (v is double || v is int || v is long || v is decimal) return Convert.ToDouble(v, Inv);
            DateTime d;
            foreach (string fmt in new string[] { "yyyy-MM-dd", "dd/MM/yyyy", "yyyy-MM-ddTHH:mm:ss" })
                if (DateTime.TryParseExact(Str(v), fmt, Inv, DateTimeStyles.None, out d)) return d.ToOADate();
            return null;
        }

        // Format-Money
        public static string Money(object v)
        {
            if (Blank(v)) return "-";
            return Convert.ToDouble(v, Inv).ToString("N2", ES) + " €";
        }
        public static string Euro0(double v) { return v.ToString("N0", ES); }

        public const double IvaRate = 0.21;
        // Split-Iva: total con IVA -> base e IVA (al céntimo). Las tasas (DGT…) son un suplido: van aparte y sin IVA
        public static void SplitIva(double total, double tasas, out double bas, out double iva, out double tx)
        {
            tx = tasas; if (tx < 0 || tx > total) tx = 0;
            double hon = total - tx;
            bas = Math.Round(hon / (1 + IvaRate), 2, MidpointRounding.AwayFromZero);
            iva = Math.Round(hon - bas, 2, MidpointRounding.AwayFromZero);
        }

        // Get-CanjeTasa: tasa DGT del tipo de permiso («B (Coche)» -> B), según Ajustes: «B=33; C=100; D=100»
        public static double? CanjeTasa(string tipo)
        {
            System.Text.RegularExpressions.Match m = PsH.M((tipo ?? "").Trim().ToUpper(), @"^([A-Z0-9+]+)");
            if (!m.Success) return null;
            string code = m.Groups[1].Value;
            foreach (string p in System.Text.RegularExpressions.Regex.Split(Config.Cur.TasasCanje ?? "", @"[;\r\n]+"))
            {
                System.Text.RegularExpressions.Match x = PsH.M(p, @"^\s*([A-Za-z0-9+]+)\s*[=:]\s*([\d.,]+)\s*€?\s*$");
                if (x.Success && x.Groups[1].Value.ToUpper() == code) return ParseMoney(x.Groups[2].Value);
            }
            return null;
        }
        // Get-TasasFor: solo el canje lleva tasas DGT
        public static double TasasFor(Tramite t, Dictionary<string, object> vals)
        {
            if (t == null || t.Id != "canje" || vals == null) return 0;
            double? x = CanjeTasa(Str(Store.Get(vals, "tipo")));
            return x.HasValue ? x.Value : 0;
        }
        // Get-Commission: comisión del colaborador en €: un % del precio o un importe fijo
        public static double Commission(Dictionary<string, object> vals, double? price)
        {
            if (vals == null) return 0;
            double? n = Num(Store.Get(vals, "comval"));
            if (!n.HasValue) return 0;
            if (Str(Store.Get(vals, "comtipo")) == "%") { if (!price.HasValue) return 0; return Math.Round(price.Value * n.Value / 100, 2, MidpointRounding.AwayFromZero); }
            return n.Value;
        }
        // Get-Margin: lo que queda: precio - tasas DGT - gastos - comisión (IVA incluido); false si no hay precio
        public static bool Margin(Tramite t, Dictionary<string, object> vals, out double margin, out double pct, out double com)
        {
            margin = 0; pct = 0; com = 0;
            double? price = Num(Store.Get(vals, "price"));
            if (!price.HasValue || price.Value <= 0) return false;
            double? g = Num(Store.Get(vals, "gastos"));
            com = Commission(vals, price);
            double m = price.Value - TasasFor(t, vals) - (g.HasValue ? g.Value : 0) - com;
            margin = Math.Round(m, 2); pct = m / price.Value * 100;
            return true;
        }

        // Sort-Object de PowerShell 5.1: powershell.exe usa la ordenación antigua de .NET (quicksort NO estable), así que los
        // empates salen en un orden concreto. Copiado de ArraySortHelper.DepthLimitedQuickSort; comprobado contra PS (n=4…40).
        public static void Sort<T>(List<T> l, Comparison<T> c)
        {
            T[] a = l.ToArray(); if (a.Length > 1) QSort(a, 0, a.Length - 1, c, 32); l.Clear(); l.AddRange(a);
        }
        static void SwapIfGreater<T>(T[] a, Comparison<T> c, int i, int j) { if (i != j && c(a[i], a[j]) > 0) { T t = a[i]; a[i] = a[j]; a[j] = t; } }
        static void QSort<T>(T[] k, int left, int right, Comparison<T> c, int depth)
        {
            do
            {
                if (depth == 0) { HSort(k, left, right, c); return; }
                int i = left, j = right, mid = i + ((j - i) >> 1);
                SwapIfGreater(k, c, i, mid); SwapIfGreater(k, c, i, j); SwapIfGreater(k, c, mid, j);
                T x = k[mid];
                do
                {
                    while (c(k[i], x) < 0) i++;
                    while (c(x, k[j]) < 0) j--;
                    if (i > j) break;
                    if (i < j) { T t = k[i]; k[i] = k[j]; k[j] = t; }
                    i++; j--;
                } while (i <= j);
                depth--;
                if (j - left <= right - i) { if (left < j) QSort(k, left, j, c, depth); left = i; }
                else { if (i < right) QSort(k, i, right, c, depth); right = j; }
            } while (left < right);
        }
        static void HSort<T>(T[] k, int lo, int hi, Comparison<T> c)
        {
            int n = hi - lo + 1;
            for (int i = n >> 1; i >= 1; i--) Down(k, i, n, lo, c);
            for (int i = n; i > 1; i--) { T t = k[lo]; k[lo] = k[lo + i - 1]; k[lo + i - 1] = t; Down(k, 1, i - 1, lo, c); }
        }
        static void Down<T>(T[] k, int i, int n, int lo, Comparison<T> c)
        {
            T d = k[lo + i - 1];
            while (i <= n >> 1)
            {
                int ch = 2 * i; if (ch < n && c(k[lo + ch - 1], k[lo + ch]) < 0) ch++;
                if (!(c(d, k[lo + ch - 1]) < 0)) break;
                k[lo + i - 1] = k[lo + ch - 1]; i = ch;
            }
            k[lo + i - 1] = d;
        }

        public static string Upper(string s) { return (s ?? "").ToUpper(Inv); }
        public static string Trim(object v) { return Str(v).Trim(); }

        // -like 'texto*' / comparación sin mayúsculas como PowerShell
        public static bool EqI(string a, string b) { return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase); }
        public static bool StartsI(string a, string b) { return (a ?? "").StartsWith(b ?? "", StringComparison.OrdinalIgnoreCase); }

        public static List<string> StrList(object o)
        {
            List<string> r = new List<string>();
            if (o == null) return r;
            if (o is string) { r.Add((string)o); return r; }
            IEnumerable en = o as IEnumerable;
            if (en == null) { r.Add(Str(o)); return r; }
            foreach (object x in en) r.Add(Str(x));
            return r;
        }
    }
}
namespace Gestem
{
    public static class PsH
    {
        // @{} de PowerShell: Hashtable sin distinguir mayúsculas (mismo orden al recorrerla que en PowerShell)
        public static System.Collections.Hashtable New() { return new System.Collections.Hashtable(System.StringComparer.CurrentCultureIgnoreCase); }
        // -match de PowerShell (sin distinguir mayúsculas)
        public static System.Text.RegularExpressions.Match M(string s, string pattern) { return System.Text.RegularExpressions.Regex.Match(s ?? "", pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase); }
        public static bool Is(string s, string pattern) { return M(s, pattern).Success; }
        // -replace de PowerShell (sin distinguir mayúsculas)
        public static string Rep(string s, string pattern, string with) { return System.Text.RegularExpressions.Regex.Replace(s ?? "", pattern, with, System.Text.RegularExpressions.RegexOptions.IgnoreCase); }
        public static bool ContainsI(System.Collections.Generic.IEnumerable<string> list, string v) { foreach (string x in list) if (string.Equals(x, v, System.StringComparison.CurrentCultureIgnoreCase)) return true; return false; }
    }
}

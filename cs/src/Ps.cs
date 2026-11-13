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
        // Split-Iva: total con IVA -> base e IVA
        public static void SplitIva(double total, out double bas, out double iva)
        {
            bas = Math.Round(total / (1 + IvaRate), 2, MidpointRounding.AwayFromZero);
            iva = Math.Round(total - bas, 2, MidpointRounding.AwayFromZero);
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

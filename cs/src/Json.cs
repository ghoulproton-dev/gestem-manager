// JSON pequeño y predecible (sin dependencias). Objetos = Dictionary<string,object> (orden de inserción),
// listas = List<object>, enteros = long, decimales = double, true/false = bool, null = null.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Gestem
{
    public static class Json
    {
        public static object Parse(string s)
        {
            if (s == null) return null;
            int i = 0;
            if (s.Length > 0 && s[0] == '﻿') i = 1;
            object v = ReadValue(s, ref i);
            SkipWs(s, ref i);
            if (i < s.Length) throw new FormatException("JSON: texto sobrante en " + i);
            return v;
        }

        static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
        }

        static object ReadValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) throw new FormatException("JSON: fin inesperado");
            char c = s[i];
            if (c == '{')
            {
                Dictionary<string, object> d = new Dictionary<string, object>();
                i++; SkipWs(s, ref i);
                if (i < s.Length && s[i] == '}') { i++; return d; }
                while (true)
                {
                    SkipWs(s, ref i);
                    string k = ReadString(s, ref i);
                    SkipWs(s, ref i);
                    if (i >= s.Length || s[i] != ':') throw new FormatException("JSON: falta ':' en " + i);
                    i++;
                    d[k] = ReadValue(s, ref i);
                    SkipWs(s, ref i);
                    if (i < s.Length && s[i] == ',') { i++; continue; }
                    if (i < s.Length && s[i] == '}') { i++; return d; }
                    throw new FormatException("JSON: objeto mal cerrado en " + i);
                }
            }
            if (c == '[')
            {
                List<object> l = new List<object>();
                i++; SkipWs(s, ref i);
                if (i < s.Length && s[i] == ']') { i++; return l; }
                while (true)
                {
                    l.Add(ReadValue(s, ref i));
                    SkipWs(s, ref i);
                    if (i < s.Length && s[i] == ',') { i++; continue; }
                    if (i < s.Length && s[i] == ']') { i++; return l; }
                    throw new FormatException("JSON: lista mal cerrada en " + i);
                }
            }
            if (c == '"') return ReadString(s, ref i);
            if (string.CompareOrdinal(s, i, "true", 0, 4) == 0) { i += 4; return true; }
            if (string.CompareOrdinal(s, i, "false", 0, 5) == 0) { i += 5; return false; }
            if (string.CompareOrdinal(s, i, "null", 0, 4) == 0) { i += 4; return null; }
            int st = i; bool frac = false;
            while (i < s.Length && "+-0123456789.eE".IndexOf(s[i]) >= 0) { if (s[i] == '.' || s[i] == 'e' || s[i] == 'E') frac = true; i++; }
            string num = s.Substring(st, i - st);
            if (num.Length == 0) throw new FormatException("JSON: valor no válido en " + st);
            if (!frac)
            {
                long l2;
                if (long.TryParse(num, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out l2)) return l2;
            }
            return double.Parse(num, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        static string ReadString(string s, ref int i)
        {
            if (s[i] != '"') throw new FormatException("JSON: se esperaba texto en " + i);
            i++;
            StringBuilder sb = new StringBuilder();
            while (i < s.Length)
            {
                char c = s[i++];
                if (c == '"') return sb.ToString();
                if (c != '\\') { sb.Append(c); continue; }
                char e = s[i++];
                switch (e)
                {
                    case '"': sb.Append('"'); break;
                    case '\\': sb.Append('\\'); break;
                    case '/': sb.Append('/'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case 'u': sb.Append((char)Convert.ToInt32(s.Substring(i, 4), 16)); i += 4; break;
                    default: throw new FormatException("JSON: escape no válido en " + i);
                }
            }
            throw new FormatException("JSON: texto sin cerrar");
        }

        public static string Write(object o) { StringBuilder sb = new StringBuilder(); WriteValue(sb, o, 0, true); return sb.ToString(); }
        public static string WriteCompact(object o) { StringBuilder sb = new StringBuilder(); WriteValue(sb, o, 0, false); return sb.ToString(); }

        static void Nl(StringBuilder sb, int ind, bool pretty) { if (!pretty) return; sb.Append("\r\n"); sb.Append(' ', ind * 2); }

        static void WriteValue(StringBuilder sb, object o, int ind, bool pretty)
        {
            if (o == null) { sb.Append("null"); return; }
            if (o is string) { WriteString(sb, (string)o); return; }
            if (o is bool) { sb.Append((bool)o ? "true" : "false"); return; }
            if (o is double || o is float || o is decimal)
            {
                double d = Convert.ToDouble(o, CultureInfo.InvariantCulture);
                if (double.IsNaN(d) || double.IsInfinity(d)) { sb.Append("null"); return; }
                sb.Append(d.ToString("R", CultureInfo.InvariantCulture)); return;
            }
            if (o is int || o is long || o is short || o is byte) { sb.Append(Convert.ToInt64(o).ToString(CultureInfo.InvariantCulture)); return; }
            if (o is DateTime) { WriteString(sb, ((DateTime)o).ToString("s", CultureInfo.InvariantCulture)); return; }
            IDictionary dict = o as IDictionary;
            if (dict != null)
            {
                sb.Append('{');
                bool first = true;
                foreach (DictionaryEntry kv in dict)
                {
                    if (!first) sb.Append(',');
                    first = false;
                    Nl(sb, ind + 1, pretty);
                    WriteString(sb, Convert.ToString(kv.Key, CultureInfo.InvariantCulture));
                    sb.Append(pretty ? ": " : ":");
                    WriteValue(sb, kv.Value, ind + 1, pretty);
                }
                if (!first) Nl(sb, ind, pretty);
                sb.Append('}');
                return;
            }
            IEnumerable en = o as IEnumerable;
            if (en != null)
            {
                sb.Append('[');
                bool first = true;
                foreach (object x in en)
                {
                    if (!first) sb.Append(',');
                    first = false;
                    Nl(sb, ind + 1, pretty);
                    WriteValue(sb, x, ind + 1, pretty);
                }
                if (!first) Nl(sb, ind, pretty);
                sb.Append(']');
                return;
            }
            WriteString(sb, Convert.ToString(o, CultureInfo.InvariantCulture));
        }

        static void WriteString(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }

        // --- ayudas para leer ---
        public static Dictionary<string, object> Obj(object o) { return o as Dictionary<string, object>; }
        public static List<object> Arr(object o) { List<object> l = o as List<object>; return l ?? new List<object>(); }
        public static object Get(object o, string key)
        {
            Dictionary<string, object> d = o as Dictionary<string, object>;
            object v;
            if (d != null && d.TryGetValue(key, out v)) return v;
            return null;
        }
        public static string Str(object o, string key) { object v = Get(o, key); return v == null ? null : Ps.Str(v); }
    }
}

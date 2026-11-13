// Trámites y campos (salen de data\tramites.json, exportado de GESTEM.ps1 tal cual).
using System;
using System.Collections.Generic;

namespace Gestem
{
    public class Field
    {
        public string Key, Col, Type, Section, Role, List, Default;
        public int Span = 1;
        public bool Required, Colored, DefaultToday, KeepCase;
        public Dictionary<string, string> Label = new Dictionary<string, string>();

        public string LabelText()
        {
            string l;
            if (Label.TryGetValue(Tx.Lang, out l) && !string.IsNullOrEmpty(l)) return l;
            Label.TryGetValue("es", out l);
            return l ?? Key;
        }
    }

    public class Tramite
    {
        public string Id, Sheet, LastCol, NetCol, IdCol, NameCol, DocCol, StatusCol, PayCol, PendCol, DaysCol, TipoCol, DefaultNationality, PriceKey;
        public int FirstRow, LastRow;
        public bool NoStats;
        public Dictionary<string, string> Title = new Dictionary<string, string>();
        public Dictionary<string, string> StatusColors = new Dictionary<string, string>();
        public Dictionary<int, string> HandCols;          // solo «tasas»: columnas fijas de la pestaña hecha a mano
        public List<Field> Fields = new List<Field>();

        public bool IsGen { get { return Id != "canje"; } }     // Test-GenTramite
        public string TitleText()
        {
            string s;
            if (Title.TryGetValue(Tx.Lang, out s) && !string.IsNullOrEmpty(s)) return s;
            Title.TryGetValue("es", out s);
            return s ?? Id;
        }
        public string TitleEs { get { string s; Title.TryGetValue("es", out s); return s ?? Id; } }
        public Field FieldByKey(string key) { foreach (Field f in Fields) if (f.Key == key) return f; return null; }
        public Field FieldByRole(string role) { foreach (Field f in Fields) if (f.Role == role) return f; return null; }
        public bool HasField(string key) { return FieldByKey(key) != null; }

        static List<Tramite> all;
        public static List<Tramite> All
        {
            get
            {
                if (all == null)
                {
                    all = new List<Tramite>();
                    foreach (object o in Json.Arr(App.ResJson("tramites.json"))) all.Add(FromJson(o));
                }
                return all;
            }
        }
        public static Tramite ById(string id) { foreach (Tramite t in All) if (t.Id == id) return t; return null; }

        static Dictionary<string, string> StrMap(object o)
        {
            Dictionary<string, string> r = new Dictionary<string, string>();
            Dictionary<string, object> d = Json.Obj(o);
            if (d != null) foreach (KeyValuePair<string, object> kv in d) r[kv.Key] = Ps.Str(kv.Value);
            return r;
        }
        static int Int(object o, int def) { object v = o; if (v == null) return def; return Convert.ToInt32(v); }
        static bool Bool(object o) { return o is bool && (bool)o; }

        static Tramite FromJson(object o)
        {
            Tramite t = new Tramite();
            t.Id = Json.Str(o, "Id"); t.Sheet = Json.Str(o, "Sheet");
            t.FirstRow = Int(Json.Get(o, "FirstRow"), 11); t.LastRow = Int(Json.Get(o, "LastRow"), 1009);
            t.LastCol = Json.Str(o, "LastCol"); t.NetCol = Json.Str(o, "NetCol"); t.IdCol = Json.Str(o, "IdCol");
            t.NameCol = Json.Str(o, "NameCol"); t.DocCol = Json.Str(o, "DocCol"); t.StatusCol = Json.Str(o, "StatusCol");
            t.PayCol = Json.Str(o, "PayCol"); t.PendCol = Json.Str(o, "PendCol"); t.DaysCol = Json.Str(o, "DaysCol");
            t.TipoCol = Json.Str(o, "TipoCol"); t.DefaultNationality = Json.Str(o, "DefaultNationality");
            t.PriceKey = Json.Str(o, "PriceKey"); t.NoStats = Bool(Json.Get(o, "NoStats"));
            t.Title = StrMap(Json.Get(o, "Title"));
            t.StatusColors = StrMap(Json.Get(o, "StatusColors"));
            Dictionary<string, object> hc = Json.Obj(Json.Get(o, "HandCols"));
            if (hc != null) { t.HandCols = new Dictionary<int, string>(); foreach (KeyValuePair<string, object> kv in hc) t.HandCols[int.Parse(kv.Key)] = Ps.Str(kv.Value); }
            foreach (object fo in Json.Arr(Json.Get(o, "Fields")))
            {
                Field f = new Field();
                f.Key = Json.Str(fo, "Key"); f.Col = Json.Str(fo, "Col"); f.Type = Json.Str(fo, "Type"); f.Section = Json.Str(fo, "Section");
                f.Role = Json.Str(fo, "Role"); f.List = Json.Str(fo, "List"); f.Default = Json.Str(fo, "Default");
                f.Span = Int(Json.Get(fo, "Span"), 1);
                f.Required = Bool(Json.Get(fo, "Required")); f.Colored = Bool(Json.Get(fo, "Colored"));
                f.DefaultToday = Bool(Json.Get(fo, "DefaultToday")); f.KeepCase = Bool(Json.Get(fo, "KeepCase"));
                f.Label = StrMap(Json.Get(fo, "Label"));
                t.Fields.Add(f);
            }
            return t;
        }
    }
}

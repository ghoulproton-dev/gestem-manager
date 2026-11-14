// config.json en %APPDATA%\GESTEM (Load-Config / Save-Config): mismas claves y mismos valores por defecto.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Gestem
{
    public class Config
    {
        public string ExcelPath = "", Language = "es", DefaultNationality = "MARROQUI", Tramite = "canje", LastUser = "";
        public bool Backup = true, OpenAfterSave = false, SumsRepaired = false, ExcelRepaired = false;
        public List<string> TabsChecked = new List<string>();
        public List<string> Colleagues = new List<string>(new string[] { "SAFAE", "CHAYMAE" });
        public string CompanyName = "", CompanyLogo = "", ExtBackupDir = "", LastExtBackup = "";
        // emisor por defecto de las facturas: la gestora
        public string InvName = "El Mahi Hmidat Chaymae", InvActivity = "Gestión de Trámites Administrativos", InvNif = "55141554C",
                      InvAddress = "Calle Los Albaladejos, 2", InvTown = "30700 Torre Pacheco (Murcia)", InvIban = "";

        public static Config Cur = new Config();

        public static void Load()
        {
            Config c = new Config();
            if (File.Exists(App.CfgPath))
            {
                try
                {
                    object j = Json.Parse(File.ReadAllText(App.CfgPath, Encoding.UTF8));
                    c.ExcelPath = S(j, "ExcelPath", c.ExcelPath); c.Language = S(j, "Language", c.Language);
                    c.DefaultNationality = S(j, "DefaultNationality", c.DefaultNationality);
                    c.Backup = B(j, "Backup", c.Backup); c.OpenAfterSave = B(j, "OpenAfterSave", c.OpenAfterSave);
                    c.Tramite = S(j, "Tramite", c.Tramite);
                    if (Json.Get(j, "TabsChecked") != null) c.TabsChecked = Ps.StrList(Json.Get(j, "TabsChecked"));
                    if (Json.Get(j, "Colleagues") != null) c.Colleagues = Ps.StrList(Json.Get(j, "Colleagues"));
                    c.LastUser = S(j, "LastUser", c.LastUser); c.CompanyName = S(j, "CompanyName", c.CompanyName);
                    c.CompanyLogo = S(j, "CompanyLogo", c.CompanyLogo); c.ExtBackupDir = S(j, "ExtBackupDir", c.ExtBackupDir);
                    c.LastExtBackup = S(j, "LastExtBackup", c.LastExtBackup);
                    c.InvName = S(j, "InvName", c.InvName); c.InvActivity = S(j, "InvActivity", c.InvActivity);
                    c.InvNif = S(j, "InvNif", c.InvNif); c.InvAddress = S(j, "InvAddress", c.InvAddress);
                    c.InvTown = S(j, "InvTown", c.InvTown); c.InvIban = S(j, "InvIban", c.InvIban);
                    c.SumsRepaired = B(j, "SumsRepaired", c.SumsRepaired); c.ExcelRepaired = B(j, "ExcelRepaired", c.ExcelRepaired);
                }
                catch (Exception ex) { App.Log("config ilegible: " + ex.Message); }
            }
            if (Array.IndexOf(Tx.Langs, c.Language) < 0) c.Language = "es";
            Cur = c;
            Tx.Lang = c.Language;
        }

        public static void Save()
        {
            Config c = Cur;
            Dictionary<string, object> o = new Dictionary<string, object>();
            o["ExcelPath"] = c.ExcelPath; o["Language"] = c.Language; o["DefaultNationality"] = c.DefaultNationality;
            o["Backup"] = c.Backup; o["OpenAfterSave"] = c.OpenAfterSave; o["Tramite"] = c.Tramite;
            o["TabsChecked"] = c.TabsChecked; o["Colleagues"] = c.Colleagues; o["LastUser"] = c.LastUser;
            o["CompanyName"] = c.CompanyName; o["CompanyLogo"] = c.CompanyLogo; o["ExtBackupDir"] = c.ExtBackupDir;
            o["LastExtBackup"] = c.LastExtBackup; o["InvName"] = c.InvName; o["InvActivity"] = c.InvActivity;
            o["InvNif"] = c.InvNif; o["InvAddress"] = c.InvAddress; o["InvTown"] = c.InvTown; o["InvIban"] = c.InvIban;
            o["SumsRepaired"] = c.SumsRepaired; o["ExcelRepaired"] = c.ExcelRepaired;
            File.WriteAllText(App.CfgPath, Json.Write(o), new UTF8Encoding(true));
        }

        // en PowerShell: if ($null -ne $j.$k) { $cfg[$k] = $j.$k }
        static string S(object j, string k, string def) { object v = Json.Get(j, k); return v == null ? def : Ps.Str(v); }
        static bool B(object j, string k, bool def)
        {
            object v = Json.Get(j, k);
            if (v == null) return def;
            if (v is bool) return (bool)v;
            if (v is string) return ((string)v).Length > 0;      // [bool]'texto' en PowerShell
            if (Ps.IsNum(v)) return Convert.ToDouble(v) != 0;
            return true;
        }
    }
}

// Modo prueba (por defecto hasta terminar el traslado): GESTEM.exe nunca toca los datos de verdad.
//  - config, registro, firebase.json y copias: cs\data\prueba (la config se copia de la de verdad la primera vez)
//  - el Excel y «GESTEM datos»: copia en cs\data\prueba\excel (la primera vez; borrar la carpeta para volver a copiar)
//  - documentos: cs\data\docs; copia externa y «nube» (OneDrive/Google Drive): cs\data\nube
// Todo lo demás lo para App.Guard. Con «GESTEM.exe --real» se usan los de verdad.
using System;
using System.IO;

namespace Gestem
{
    public static class TestMode
    {
        public static string Root { get { return Path.Combine(App.AppDir, "data", "prueba"); } }
        public static string DocsRoot { get { return Path.Combine(App.AppDir, "data", "docs"); } }
        public static string CloudRoot { get { return Path.Combine(App.AppDir, "data", "nube"); } }    // hace de OneDrive / Google Drive
        public static string Banner = "MODO PRUEBA (copia de los datos; GESTEM.exe --real para los de verdad)";

        public static void Setup()
        {
            string realCfg = App.CfgPath;
            App.AddSafeRoot(Root); App.AddSafeRoot(DocsRoot); App.AddSafeRoot(CloudRoot);
            Directory.CreateDirectory(CloudRoot); App.CloudRootsOverride = new System.Collections.Generic.List<string> { CloudRoot };
            App.DataDir = Root; App.CfgPath = Path.Combine(Root, "config.json"); App.BackupDir = Path.Combine(Root, "copias");
            App.LogPath = Path.Combine(Root, "gestem.log"); FbState.Path = Path.Combine(Root, "firebase.json");
            Directory.CreateDirectory(Root); Directory.CreateDirectory(DocsRoot);
            if (!File.Exists(App.CfgPath) && File.Exists(realCfg)) File.Copy(realCfg, App.CfgPath);    // solo se lee la de verdad
            string realAdm = Path.Combine(Path.GetDirectoryName(realCfg), "admin.dat"), adm = Path.Combine(Root, "admin.dat");
            if (!File.Exists(adm) && File.Exists(realAdm)) { File.Copy(realAdm, adm); }    // la contraseña de administrador de este PC
            Config.Load();
            string x = Ps.Str(Config.Cur.ExcelPath);
            if (x.Length > 0 && !App.IsSafe(x))
            {
                string dir = Path.Combine(Root, "excel"), copy = Path.Combine(dir, Path.GetFileName(x));
                if (!File.Exists(copy) && File.Exists(x))
                {
                    Directory.CreateDirectory(Path.Combine(dir, "GESTEM datos"));
                    string sd = Path.Combine(Path.GetDirectoryName(x), "GESTEM datos");
                    if (Directory.Exists(sd)) foreach (string f in Directory.GetFiles(sd, "*.json")) File.Copy(f, Path.Combine(dir, "GESTEM datos", Path.GetFileName(f)));
                    File.Copy(x, copy);
                    App.Log("modo prueba: copia de " + x + " y de sus datos en " + dir);
                }
                Config.Cur.ExcelPath = File.Exists(copy) ? copy : "";
            }
            Config.Cur.DocsDir = DocsRoot;
            if (Ps.Str(Config.Cur.ExtBackupDir).Length > 0 && !App.IsSafe(Config.Cur.ExtBackupDir)) Config.Cur.ExtBackupDir = Path.Combine(CloudRoot, "GESTEM copias");    // copia externa: en la «nube» de prueba
            Config.Cur.FbSync = false;
            Config.Save();
        }
    }
}
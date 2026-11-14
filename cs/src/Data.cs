// Cargar y guardar: GESTEM primero, el Excel después. Copiado de Read-ExcelMeta / Import-TramiteFromExcel /
// Update-StoreMeta / Save-ClientStore / Sync-RecordToExcel / Save-PendingBatchToFile / Repair-TotalsFormulas /
// Sync-PendingToExcel / Refresh-FromExcel / Load-Data / Get-CachedData / Sync-DataBackup.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Gestem
{
    public class SaveResult { public string Id; public bool Synced; }

    public static class Data
    {
        public static void UseData(TramiteData d)
        {
            G.Lists = d.Lists; G.Tariff = d.Tariff; G.DefaultPrice = d.DefaultPrice;
            G.Rows = d.Rows; G.FirstEmpty = d.FirstEmpty; G.NextId = d.NextId;
            if (G.OnUseData != null) G.OnUseData(d);
        }

        static Meta ReadExcelMeta(Book h, Tramite t, TramiteData d)
        {
            object ws = ExcelCom.GetSheet(h, false);
            Meta m = new Meta();
            foreach (KeyValuePair<string, string[]> kv in d.Lists) if (kv.Key.StartsWith("excel:", StringComparison.OrdinalIgnoreCase)) m.Lists[kv.Key] = kv.Value;
            if (t.IsGen)
            {
                // pestaña sin listas desplegables (p. ej. creada por GESTEM): listas de siempre
                if (!HasText(m.Lists, "excel:N")) m.Lists["excel:N"] = Hand.GenStatusList;
                if (!HasText(m.Lists, "excel:M")) m.Lists["excel:M"] = new string[] { "Pagado", "Parcial", "Pendiente" };
            }
            if (!string.IsNullOrEmpty(t.TipoCol)) m.TariffNet = ExcelCom.Run<Dictionary<string, double>>(delegate { return ExcelCom.ReadTariffNet(ws, t.TipoCol); });
            m.Tariff = d.Tariff; m.DefaultPrice = d.DefaultPrice;
            string first = Com.S(ws, t.IdCol + t.FirstRow);
            Match mm = PsH.M(first, @"^(.*?)(\d+)$");
            if (mm.Success) { m.IdPrefix = mm.Groups[1].Value; m.IdWidth = mm.Groups[2].Value.Length; }
            else if (first.Trim().Length == 0 && t.IsGen) m.IdPrefix = Hand.DefaultIdPrefix(t);
            return m;
        }
        static bool HasText(Dictionary<string, string[]> l, string k)
        {
            if (!l.ContainsKey(k) || l[k] == null) return false;
            foreach (string s in l[k]) if ((s ?? "").Trim().Length > 0) return true;
            return false;
        }

        static double? PriceFromBook(Book h, Tramite t)
        {
            if (h == null || string.IsNullOrEmpty(t.PriceKey)) return null;
            try { return ExcelCom.ReadTramitePrice(h.Wb, t.PriceKey); } catch { return null; }
        }

        public static StoreData ImportTramiteFromExcel(Book h, Tramite t)
        {
            if (t.IsGen && ExcelCom.Run<object>(delegate { return ExcelCom.FindTramiteSheet(h.Wb, t); }) == null)
            {
                // el Excel no tiene esta pestaña: se empieza vacío y la pestaña se crea al guardar el primer cliente
                StoreData st = Hand.NewEmptyStore(t, PriceFromBook(h, t));
                try { StoreData old = Store.Read(t); if (old != null) st.Adopted = new List<string>(old.Adopted); } catch { }
                Store.Write(t, st);
                App.Log(string.Format("sin pestaña en el Excel, trámite empezado vacío: {0}", t.Sheet));
                return Store.Read(t);
            }
            TramiteData d = ExcelCom.ReadTramiteData(h);
            Meta meta = ReadExcelMeta(h, t, d);
            StoreData store = new StoreData(); store.Meta = meta;
            foreach (ClientRow r in d.Rows)
            {
                ClientRec c = new ClientRec();
                foreach (Field f in t.Fields) c.F[f.Key] = Store.Get(r.Data, f.Key);
                c.Row = r.Row; c.Id = r.Id; c.Sync = true; c.Updated = DateTime.Now.ToString("s", Ps.Inv); c.By = "Excel";
                store.Clients.Add(c);
            }
            try { StoreData old = Store.Read(t); if (old != null) store.Adopted = new List<string>(old.Adopted); } catch { }
            Store.Write(t, store);
            App.Log(string.Format("importó del Excel: {0} ({1} clientes)", t.Sheet, store.Clients.Count));
            return Store.Read(t);     // mismos valores que se guardaron (fechas e importes ya convertidos)
        }

        public static void UpdateStoreMeta(Book h, Tramite t)
        {
            TramiteData d = ExcelCom.ReadTramiteData(h);
            Meta meta = ReadExcelMeta(h, t, d);
            string lk = Store.EnterLock();
            try
            {
                StoreData s = Store.Read(t);
                if (s == null) return;
                s.Meta = meta;
                Store.Write(t, s);
            }
            finally { Store.ExitLock(lk); }
        }

        static ClientRec FindRow(StoreData s, int row) { foreach (ClientRec c in s.Clients) if (c.Row == row) return c; return null; }

        // Save-ClientStore: guardar primero en GESTEM, luego en el Excel
        public static SaveResult SaveClientStore(Dictionary<string, object> vals, int? editRow)
        {
            Tramite t = G.T;
            string lk = Store.EnterLock();
            ClientRec rec = null;
            try
            {
                StoreData store = Store.Read(t);
                if (store == null) throw new GestemException(Tx.L("errNoStore"));
                Dictionary<string, object> FV = new Dictionary<string, object>();
                foreach (Field f in t.Fields) FV[f.Key] = Store.Get(vals, f.Key);
                FV = Store.Resolve(t, FV, store.Meta);
                if (editRow.HasValue && editRow.Value >= Hand.RowBase)
                {
                    // cliente de una pestaña hecha a mano editado en GESTEM: pasa a ser de GESTEM (fila nueva)
                    ClientRec old = FindRow(store, editRow.Value);
                    if (old != null)
                    {
                        store.Clients.Remove(old);
                        List<string> ad = new List<string>();
                        foreach (string a in store.Adopted) if (a.Trim().Length > 0 && !ad.Contains(a)) ad.Add(a);
                        string nm = Ps.Str(Store.Get(old.F, "name"));
                        if (nm.Trim().Length > 0 && !ad.Contains(nm)) ad.Add(nm);
                        store.Adopted = ad;
                        App.Log(string.Format("cliente de la pestaña hecha a mano pasa a GESTEM: {0}", nm));
                    }
                    editRow = null;
                }
                if (editRow.HasValue && editRow.Value != 0) rec = FindRow(store, editRow.Value);
                if (rec != null) rec.F = FV;
                else
                {
                    HashSet<int> used = new HashSet<int>(); foreach (ClientRec c in store.Clients) used.Add(c.Row);
                    int? row = null;
                    for (int r = t.FirstRow; r <= t.LastRow; r++) if (!used.Contains(r)) { row = r; break; }
                    if (!row.HasValue) throw new GestemException(Tx.L("errFull"));
                    rec = new ClientRec(); rec.Row = row.Value; rec.Id = Store.RowId(t, store.Meta, row.Value); rec.F = FV; rec.New = true;
                    store.Clients.Add(rec);
                }
                rec.Sync = false;
                rec.Updated = DateTime.Now.ToString("s", Ps.Inv);
                rec.By = App.User.Length > 0 ? App.User : "ADMIN";
                Store.Write(t, store);
            }
            finally { Store.ExitLock(lk); }
            bool synced = SyncRecordToExcel(t, rec, null, false);
            SaveResult sr = new SaveResult(); sr.Id = rec.Id; sr.Synced = synced;
            return sr;
        }

        public static bool SyncRecordToExcel(Tramite t, ClientRec rec, Book h, bool fileOnly)
        {
            Tramite prevT = G.T; Dictionary<string, double> prevTar = G.Tariff;
            try
            {
                G.T = t;
                StoreData s0 = Store.Read(t);
                if (s0 != null) G.Tariff = s0.Meta.Tariff;
                Dictionary<string, object> vals = new Dictionary<string, object>(rec.F);
                FileResult r = ExcelCom.SaveClient(vals, rec.Row, Ps.Str(Store.Get(rec.F, "name")), rec.New, h, fileOnly);
                if (r == null) return false;     // fileOnly: no se pudo en el archivo, lo hará Excel después
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    ClientRec c = FindRow(s, rec.Row);
                    if (c != null)
                    {
                        if (r.Row != rec.Row) { c.Row = r.Row; App.Log(string.Format("expediente {0} pasa a la fila {1} del Excel", rec.Id, r.Row)); }
                        if (r.Id.Length > 0) c.Id = r.Id;
                        c.Sync = true; c.New = false;
                        Store.Write(t, s);
                    }
                    rec.Row = r.Row; if (r.Id.Length > 0) rec.Id = r.Id;
                }
                finally { Store.ExitLock(lk); }
                return true;
            }
            catch (Exception e)
            {
                App.Log(string.Format("Excel no actualizado ({0}): {1}", rec.Id, e.Message));
                return false;
            }
            finally { G.T = prevT; G.Tariff = prevTar; }
        }

        public static bool SavePendingBatchToFile(Tramite t, List<ClientRec> recs)
        {
            Tramite prevT = G.T; Dictionary<string, double> prevTar = G.Tariff;
            try
            {
                G.T = t;
                StoreData s0 = Store.Read(t); if (s0 != null) G.Tariff = s0.Meta.Tariff;
                List<FileItem> items = new List<FileItem>();
                foreach (ClientRec rec in recs)
                {
                    FileItem it = new FileItem(); it.Ops = ExcelFile.ClientOps(t, new Dictionary<string, object>(rec.F), G.Tariff);
                    it.Row = rec.Row; it.Expect = Ps.Str(Store.Get(rec.F, "name")); it.IsNew = rec.New;
                    items.Add(it);
                }
                List<FileResult> res = null;
                try { res = ExcelFile.SaveMany(t, items, false); } catch (Exception e) { App.Log("Excel directo no posible: " + e.Message); }
                if (res == null || res.Count == 0) return false;
                string lk = Store.EnterLock();
                try
                {
                    StoreData s = Store.Read(t);
                    for (int i = 0; i < recs.Count; i++)
                    {
                        ClientRec rec = recs[i]; FileResult r = res[i];
                        ClientRec c = FindRow(s, rec.Row);
                        if (c == null) continue;
                        if (r.Row != rec.Row) { c.Row = r.Row; App.Log(string.Format("expediente {0} pasa a la fila {1} del Excel", rec.Id, r.Row)); }
                        if (r.Id.Length > 0) c.Id = r.Id;
                        c.Sync = true; c.New = false;
                    }
                    Store.Write(t, s);
                }
                finally { Store.ExitLock(lk); }
                return true;
            }
            finally { G.T = prevT; G.Tariff = prevTar; }
        }

        // una vez por PC: quitar las SUM(...) que versiones anteriores copiaron en filas de clientes
        public static void RepairTotalsFormulas()
        {
            if (Config.Cur.SumsRepaired) return;
            bool ok = true; Tramite prevT = G.T; Dictionary<string, double> prevTar = G.Tariff;
            try
            {
                foreach (Tramite t in Tramite.All)
                {
                    StoreData s = null; try { s = Store.Read(t); } catch { }
                    if (s == null) continue;
                    List<ClientRec> recs = new List<ClientRec>(); foreach (ClientRec c in s.Clients) if (c.Row < Hand.RowBase && c.Sync) recs.Add(c);
                    if (recs.Count == 0) continue;
                    G.T = t; G.Tariff = s.Meta.Tariff;
                    List<FileItem> items = new List<FileItem>();
                    foreach (ClientRec rec in recs)
                    {
                        FileItem it = new FileItem(); it.Ops = ExcelFile.ClientOps(t, new Dictionary<string, object>(rec.F), G.Tariff);
                        it.Row = rec.Row; it.Expect = Ps.Str(Store.Get(rec.F, "name")); it.IsNew = false; items.Add(it);
                    }
                    List<FileResult> res = null;
                    try { res = ExcelFile.SaveMany(t, items, true); } catch (Exception e) { App.Log("reparar totales (" + t.Sheet + "): " + e.Message); }
                    if ((res == null || res.Count == 0) && t.IsGen) continue;    // trámite sin pestaña de GESTEM
                    if (res == null || res.Count == 0) ok = false;
                }
            }
            finally { G.T = prevT; G.Tariff = prevTar; }
            if (ok) { Config.Cur.SumsRepaired = true; Config.Save(); }
        }

        // clientes guardados en GESTEM que todavía no están en el Excel
        public static int SyncPendingToExcel()
        {
            List<KeyValuePair<Tramite, ClientRec>> todo = new List<KeyValuePair<Tramite, ClientRec>>();
            foreach (Tramite t in Tramite.All)
            {
                StoreData s = null; try { s = Store.Read(t); } catch { }
                if (s != null) foreach (ClientRec c in s.Clients) if (!c.Sync) todo.Add(new KeyValuePair<Tramite, ClientRec>(t, c));
            }
            if (todo.Count == 0) return 0;
            Book h = null; int n = 0;
            List<KeyValuePair<Tramite, ClientRec>> left = new List<KeyValuePair<Tramite, ClientRec>>();
            foreach (Tramite t in Tramite.All)
            {
                List<KeyValuePair<Tramite, ClientRec>> grp = todo.FindAll(delegate(KeyValuePair<Tramite, ClientRec> x) { return x.Key.Id == t.Id; });
                if (grp.Count == 0) continue;
                List<ClientRec> recs = new List<ClientRec>(); foreach (KeyValuePair<Tramite, ClientRec> x in grp) recs.Add(x.Value);
                if (SavePendingBatchToFile(t, recs)) n += grp.Count; else left.AddRange(grp);
            }
            if (left.Count == 0) { if (n > 0) App.Log(string.Format("Excel actualizado con {0} cliente(s) pendiente(s)", n)); return n; }
            try
            {
                h = ExcelCom.Run<Book>(delegate { return ExcelCom.Open(true); });
                if (Config.Cur.Backup) ExcelFile.BackupExcel();
                foreach (KeyValuePair<Tramite, ClientRec> x in left) if (SyncRecordToExcel(x.Key, x.Value, h, false)) n++;
                Book hh = h; ExcelCom.Run(delegate { ExcelCom.Close(hh, true); }); h = null;
            }
            catch (Exception e)
            {
                App.Log("sincronizar con el Excel: " + e.Message);
                if (h != null) { try { ExcelCom.Close(h, false); } catch { } }
            }
            if (n > 0) App.Log(string.Format("Excel actualizado con {0} cliente(s) pendiente(s)", n));
            return n;
        }

        public static int PendingSyncCount()
        {
            int n = 0;
            if (G.Cache == null) return 0;
            foreach (Tramite t in Tramite.All) if (G.Cache.ContainsKey(t.Id)) foreach (ClientRow r in G.Cache[t.Id].Rows) if (!r.Sync) n++;
            return n;
        }

        // «Actualizar»: listas leídas del archivo; Excel solo para los trámites cuyas listas están en otras celdas
        public static bool RefreshFromExcel()
        {
            Book h = null; Tramite cur = G.T;
            try
            {
                List<Tramite> viaExcel = new List<Tramite>(Tramite.All);
                try
                {
                    System.Collections.Hashtable dv = Xlsx.ReadValidationLists(Config.Cur.ExcelPath);
                    List<string> done = new List<string>();
                    viaExcel = new List<Tramite>();
                    foreach (Tramite t in Tramite.All)
                    {
                        string res = "excel";
                        try { res = Hand.UpdateListsFromFile(t, dv); } catch (Exception e) { App.Log("listas sin Excel (" + t.Sheet + "): " + e.Message); }
                        if (res == "ok") done.Add(t.Id);
                        if (res == "excel") viaExcel.Add(t);
                    }
                    List<string> ve = new List<string>(); foreach (Tramite t in viaExcel) ve.Add(t.Id);
                    App.Log(string.Format("actualizar: {0} pestañas leídas del archivo · listas sin Excel: {1} · con Excel: {2}", dv.Count, string.Join(", ", done.ToArray()), string.Join(", ", ve.ToArray())));
                }
                catch (Exception e) { App.Log("listas sin Excel: " + e.Message); }
                if (viaExcel.Count > 0 && ExcelCom.Installed())
                {
                    h = ExcelCom.Run<Book>(delegate { return ExcelCom.Open(false); });
                    foreach (Tramite t in viaExcel)
                    {
                        G.T = t;
                        try { UpdateStoreMeta(h, t); } catch (Exception e) { App.Log("listas (" + t.Sheet + "): " + e.Message); }
                    }
                    G.T = cur;
                    ExcelCom.Close(h, false); h = null;
                }
            }
            catch (Exception e)
            {
                G.T = cur;
                if (h != null) { try { ExcelCom.Close(h, false); } catch { } }
                App.Log("actualizar desde el Excel: " + e.Message);
            }
            try { Hand.SyncTabs(true); } catch (Exception e) { App.Log("pestañas propias: " + e.Message); }
            SyncPendingToExcel();
            return LoadData(true, false);
        }

        // Load-Data: lee los datos de GESTEM (rápido, sin Excel); la primera vez los importa del Excel
        public static bool LoadData(bool all, bool quiet)
        {
            if (!quiet) G.Status(Tx.L("loading"));
            Book h = null;
            Tramite cur = G.T;
            try
            {
                if (string.IsNullOrEmpty(Config.Cur.ExcelPath) || !File.Exists(Config.Cur.ExcelPath)) throw new GestemException(Tx.L("errNoFile"));
                Dictionary<string, TramiteData> newCache = G.Cache;
                if (all || newCache == null) newCache = new Dictionary<string, TramiteData>();
                List<Tramite> list = new List<Tramite>(); list.Add(cur);
                if (all) foreach (Tramite t in Tramite.All) if (t.Id != cur.Id) list.Add(t);
                foreach (Tramite t in list)
                {
                    G.T = t;
                    try
                    {
                        StoreData s = Store.Read(t);
                        if (s != null && !quiet && t.IsGen && !Config.Cur.TabsChecked.Contains(t.Id))
                        {
                            // una sola vez por trámite: ¿el Excel tiene una pestaña con el formato de GESTEM?
                            bool done = false;
                            try
                            {
                                if (h == null) h = ExcelCom.Run<Book>(delegate { return ExcelCom.Open(false); });
                                Book hh = h;
                                object tab = ExcelCom.Run<object>(delegate { return ExcelCom.FindTramiteSheet(hh.Wb, t); });
                                string lk = Store.EnterLock();
                                try
                                {
                                    StoreData s2 = Store.Read(t);
                                    if (s2 != null && tab != null && s2.Clients.Count == 0) s = ImportTramiteFromExcel(h, t);
                                    else if (s2 != null && tab == null)
                                    {
                                        List<ClientRec> keep = new List<ClientRec>();
                                        foreach (ClientRec c in s2.Clients) if (c.By != "Excel" || c.Row >= Hand.RowBase) keep.Add(c);
                                        if (keep.Count != s2.Clients.Count)
                                        {
                                            App.Log(string.Format("{0}: {1} cliente(s) leídos de una pestaña que no es de GESTEM, quitados", t.Sheet, s2.Clients.Count - keep.Count));
                                            s2.Clients = keep; s2.Meta = Hand.NewEmptyStore(t, PriceFromBook(h, t)).Meta;
                                            Store.Write(t, s2); s = Store.Read(t);
                                        }
                                    }
                                }
                                finally { Store.ExitLock(lk); }
                                done = true;
                            }
                            catch (Exception e) { App.Log("pestaña no leída (" + t.Sheet + "): " + e.Message); }
                            finally
                            {
                                if (done)
                                {
                                    List<string> tc = new List<string>();
                                    foreach (string x in Config.Cur.TabsChecked) if (x.Length > 0 && !tc.Contains(x)) tc.Add(x);
                                    if (!tc.Contains(t.Id)) tc.Add(t.Id);
                                    Config.Cur.TabsChecked = tc; Config.Save();
                                }
                            }
                        }
                        if (s == null && quiet) { if (G.Cache != null && G.Cache.ContainsKey(t.Id)) newCache[t.Id] = G.Cache[t.Id]; continue; }
                        if (s == null && ExcelCom.TabKnownMissing(t))
                        {
                            if (t.Id != cur.Id) continue;
                            throw new GestemException(string.Format(Tx.L("errNoSheet"), t.Sheet));
                        }
                        if (s == null)
                        {
                            if (h == null) h = ExcelCom.Run<Book>(delegate { return ExcelCom.Open(false); });
                            string lk = Store.EnterLock();
                            try { s = Store.Read(t); if (s == null) s = ImportTramiteFromExcel(h, t); } finally { Store.ExitLock(lk); }
                        }
                        newCache[t.Id] = Store.BuildData(t, s);
                    }
                    catch (Exception e)
                    {
                        if (t.Id == cur.Id) throw;
                        App.Log("pestaña no leída (" + t.Sheet + "): " + e.Message);
                    }
                }
                G.T = cur;
                if (h != null) { ExcelCom.Close(h, false); h = null; }
                if (!newCache.ContainsKey(cur.Id)) throw new GestemException(Tx.L("errNoStore"));
                G.Cache = newCache;
                G.CacheStamp = Store.FileStamp();
                UseData(G.Cache[cur.Id]);
                return true;
            }
            catch (Exception e)
            {
                G.T = cur;
                if (h != null) { try { ExcelCom.Close(h, false); } catch { } }
                App.Log("Load-Data: " + e.Message);
                if (!quiet)
                {
                    G.Status(e.Message);
                    // el aviso de pestaña que falta sale una sola vez (no en cada cambio de trámite)
                    if (!(G.MissingTabs.ContainsKey(cur.Id) && G.MissingWarned.ContainsKey(cur.Id)))
                    {
                        if (G.MissingTabs.ContainsKey(cur.Id)) G.MissingWarned[cur.Id] = true;
                        G.Error(e.Message);
                    }
                }
                return false;
            }
        }

        // cambiar de trámite: instantáneo si los datos no han cambiado desde la última lectura
        public static bool CachedData()
        {
            if (G.Cache != null && G.Cache.ContainsKey(G.T.Id) && G.CacheStamp == Store.FileStamp()) { UseData(G.Cache[G.T.Id]); return true; }
            return LoadData(true, false);
        }

        // copia de los datos de GESTEM en la carpeta externa
        public static void SyncDataBackup(string dir)
        {
            string root = Store.Dir();
            if (root == null || !Directory.Exists(root)) return;
            string dst = Path.Combine(dir, "datos GESTEM");
            if (!Directory.Exists(dst)) Directory.CreateDirectory(dst);
            foreach (string f in Directory.GetFiles(root, "*.json")) { try { File.Copy(f, Path.Combine(dst, Path.GetFileName(f)), true); } catch { } }
        }
    }
}

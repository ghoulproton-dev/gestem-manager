// Etapa 4a: Resumen (cifras por trámite o por estado, tarjetas y gráfica). Copiado de Fill-StatsFilters, New-StatCell,
// Draw-ChartLine, Update-Chart, Start-CountUp y Update-Stats de GESTEM.ps1 (y sus eventos).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Gestem
{
    public class StatAgg { public int N; public double Bill, Paid, Pend, Net; public int Debt; }
    public class ChartBucket { public DateTime From, To; public string Label, Long; public int N; public double Bill, Paid; }

    public partial class MainWin
    {
        public string StatsScopeId;            // $script:StatsScopeId
        bool statsFilling, animateChart;
        public List<ChartBucket> ChartBuckets;
        DispatcherTimer countTimer;

        static string Euro0(double v) { return v.ToString("N0", Ps.ES) + " €"; }
        static SolidColorBrush Br(string c) { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(c)); }
        bool StatsVisible { get { return U("PanelStats").Visibility == Visibility.Visible; } }
        string StatsWho() { ComboBox w = U<ComboBox>("StatsWho"); return w.SelectedIndex > 0 ? Ps.Str(w.SelectedItem) : null; }

        void WireStats()
        {
            foreach (string n in new string[] { "ChartDay", "ChartWeek", "ChartMonth" }) U<RadioButton>(n).Checked += delegate { if (StatsVisible) UpdateChart(); };
            U<Canvas>("ChartCanvas").SizeChanged += delegate { DrawChartLine(); };
            U<ComboBox>("StatsScope").SelectionChanged += delegate
            {
                if (statsFilling) return;
                int ix = U<ComboBox>("StatsScope").SelectedIndex;
                StatsScopeId = ix > 0 ? Tramite.All[ix - 1].Id : null;
                if (StatsVisible) UpdateStats();
            };
            U<Button>("StatsAllBtn").Click += delegate { StatsScopeId = null; FillStatsFilters(); UpdateStats(); };
            U<ComboBox>("StatsWho").SelectionChanged += delegate { if (StatsVisible) UpdateStats(); };
            U<ComboBox>("StatsPeriod").SelectionChanged += delegate { if (StatsVisible) UpdateStats(); };
            U<Button>("StatsRefresh").Click += delegate { Win.Cursor = Cursors.Wait; Data.LoadData(true, false); FillStatsFilters(); UpdateStats(); Win.Cursor = null; };
        }

        public void SetStatsTexts()
        {
            U<ContentControl>("ChartDay").Content = L("chDay"); U<ContentControl>("ChartWeek").Content = L("chWeek"); U<ContentControl>("ChartMonth").Content = L("chMonth");
            U<TextBlock>("ChartLegPaid").Text = L("chPaid"); U<TextBlock>("ChartLegPend").Text = L("chPend"); U<TextBlock>("ChartLegMoney").Text = L("chMoney");
            U<TextBlock>("StatsTitle").Text = L("statsTitle"); U<TextBlock>("StatsTable").Text = L("statsTable");
            U<TextBlock>("KpiBillL").Text = L("kpiBill"); U<TextBlock>("KpiPaidL").Text = L("kpiPaid"); U<TextBlock>("KpiPendL").Text = L("kpiPend"); U<TextBlock>("KpiNetL").Text = L("kpiNet");
            FillStatsFilters();
        }

        // $script:Cache es un @{}: se recorre en el orden de un Hashtable de PowerShell
        IEnumerable<TramiteData> CacheValuesPs()
        {
            List<TramiteData> l = new List<TramiteData>();
            if (G.Cache == null) return l;
            Hashtable h = PsH.New();
            foreach (Tramite t in Tramite.All) if (G.Cache.ContainsKey(t.Id)) h[t.Id] = G.Cache[t.Id];
            foreach (DictionaryEntry e in h) l.Add((TramiteData)e.Value);
            return l;
        }

        public void FillStatsFilters()
        {
            statsFilling = true;
            ComboBox who = U<ComboBox>("StatsWho"), scope = U<ComboBox>("StatsScope"), per = U<ComboBox>("StatsPeriod");
            int whoSel = who.SelectedIndex, perSel = per.SelectedIndex;
            who.Items.Clear();
            who.Items.Add(L("whoAll"));
            List<string> names = new List<string>();
            foreach (string c in Config.Cur.Colleagues) if (Ps.Str(c).Trim().Length > 0 && !names.Contains(c)) names.Add(c);
            foreach (TramiteData d in CacheValuesPs()) foreach (ClientRow r in d.Rows) { string x = Ps.Str(r.Resp).Trim().ToUpper(); if (x.Length > 0 && !names.Contains(x)) names.Add(x); }
            foreach (string n in names) who.Items.Add(n);
            scope.Items.Clear();
            scope.Items.Add(L("scopeAll"));
            int si = 0, k = 0;
            foreach (Tramite t in Tramite.All)
            {
                k++;
                scope.Items.Add(t.TitleText());
                if (StatsScopeId != null && t.Id == StatsScopeId) si = k;
            }
            scope.SelectedIndex = si;
            per.Items.Clear();
            foreach (string p in new string[] { "perAll", "perMonth", "perPrev", "perYear" }) per.Items.Add(L(p));
            who.SelectedIndex = whoSel >= 0 && whoSel < who.Items.Count ? whoSel : 0;
            per.SelectedIndex = perSel >= 0 ? perSel : 0;
            statsFilling = false;
        }

        static TextBlock StatCell(string text, int col, TextAlignment align, string color, bool bold, double size)
        {
            TextBlock t = new TextBlock();
            t.Text = text; t.FontSize = size; t.Foreground = Br(color);
            t.VerticalAlignment = VerticalAlignment.Center; t.TextAlignment = align; t.Margin = new Thickness(8, 0, 8, 0);
            if (bold) t.FontWeight = FontWeights.SemiBold;
            Grid.SetColumn(t, col);
            return t;
        }

        void DrawChartLine()
        {
            Canvas cv = U<Canvas>("ChartCanvas");
            cv.Children.Clear();
            List<ChartBucket> bk = ChartBuckets;
            if (bk == null || bk.Count < 2) return;
            double W = cv.ActualWidth, H = cv.ActualHeight;
            if (W <= 0 || H <= 0) return;
            int n = bk.Count;
            double max = 1; foreach (ChartBucket b in bk) if (b.N > max) max = b.N;
            double top = 30, bottom = 12, plotH = H - top - bottom, step = W / n;
            // líneas guía
            foreach (double f in new double[] { 0, 0.5, 1 })
            {
                double y = top + plotH * f;
                Line gl = new Line(); gl.X1 = 0; gl.X2 = W; gl.Y1 = y; gl.Y2 = y;
                gl.Stroke = Br(f == 1 ? "#E4DFE0" : "#F3EFEF"); gl.StrokeThickness = 1;
                if (f < 1) gl.StrokeDashArray = new DoubleCollection(new double[] { 4, 4 });
                cv.Children.Add(gl);
            }
            List<Point> pts = new List<Point>();
            for (int i = 0; i < n; i++) pts.Add(new Point(step * (i + 0.5), top + plotH * (1 - bk[i].N / max)));
            // línea recta entre puntos y área sombreada debajo
            double bas = top + plotH;
            PathFigure af = new PathFigure(); af.StartPoint = new Point(pts[0].X, bas); af.IsClosed = true;
            af.Segments.Add(new LineSegment(pts[0], false));
            for (int i = 1; i < n; i++) af.Segments.Add(new LineSegment(pts[i], false));
            af.Segments.Add(new LineSegment(new Point(pts[n - 1].X, bas), false));
            PathGeometry ageo = new PathGeometry(); ageo.Figures.Add(af);
            Path area = new Path(); area.Data = ageo;
            LinearGradientBrush ag = new LinearGradientBrush(); ag.StartPoint = new Point(0, 0); ag.EndPoint = new Point(0, 1);
            ag.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#2A9E1820"), 0));
            ag.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#009E1820"), 1));
            area.Fill = ag; cv.Children.Add(area);
            // periodo actual: línea vertical fina
            double vx = pts[n - 1].X;
            Line vl = new Line(); vl.X1 = vx; vl.X2 = vx; vl.Y1 = pts[n - 1].Y; vl.Y2 = bas;
            vl.Stroke = Br("#559E1820"); vl.StrokeThickness = 1; vl.StrokeDashArray = new DoubleCollection(new double[] { 3, 3 });
            cv.Children.Add(vl);
            // tramos: verde si sube, rojo si baja, gris si igual
            List<string> colors = new List<string>(); colors.Add("#A69C9E");
            for (int i = 1; i < n; i++)
            {
                string c = bk[i].N > bk[i - 1].N ? "#1F9D55" : bk[i].N < bk[i - 1].N ? "#D8202A" : "#B8B0B2";
                colors.Add(c);
                PathFigure fig = new PathFigure(); fig.StartPoint = pts[i - 1]; fig.Segments.Add(new LineSegment(pts[i], true));
                PathGeometry geo = new PathGeometry(); geo.Figures.Add(fig);
                Path seg = new Path(); seg.Data = geo; seg.Stroke = Br(c); seg.StrokeThickness = 2.6;
                seg.StrokeStartLineCap = PenLineCap.Round; seg.StrokeEndLineCap = PenLineCap.Round; seg.StrokeLineJoin = PenLineJoin.Round;
                cv.Children.Add(seg);
            }
            // puntos y cifras
            for (int i = 0; i < n; i++)
            {
                bool isNow = i == n - 1;
                double d = isNow ? 13 : 8;
                if (isNow)
                {
                    Ellipse halo = new Ellipse(); halo.Width = 26; halo.Height = 26;
                    halo.Fill = Br(colors[i] == "#1F9D55" ? "#261F9D55" : colors[i] == "#D8202A" ? "#26D8202A" : "#26A69C9E");
                    Canvas.SetLeft(halo, pts[i].X - 13); Canvas.SetTop(halo, pts[i].Y - 13); cv.Children.Add(halo);
                }
                Ellipse dot = new Ellipse(); dot.Width = d; dot.Height = d;
                dot.Fill = isNow ? Br(colors[i]) : Br("#FFFFFF");
                dot.Stroke = isNow ? Br("#FFFFFF") : Br(colors[i]); dot.StrokeThickness = isNow ? 2.5 : 2;
                dot.ToolTip = string.Format(L("chTip"), bk[i].Long, bk[i].N, Euro0(bk[i].Bill), Euro0(bk[i].Paid));
                dot.Cursor = Cursors.Hand;
                Canvas.SetLeft(dot, pts[i].X - d / 2); Canvas.SetTop(dot, pts[i].Y - d / 2); cv.Children.Add(dot);
                if (bk[i].N > 0 || isNow)
                {
                    TextBlock tx = new TextBlock(); tx.Text = bk[i].N.ToString(Ps.Inv); tx.Width = 50; tx.TextAlignment = TextAlignment.Center;
                    tx.FontFamily = new FontFamily("Segoe UI Variable Display, Segoe UI");
                    tx.FontSize = isNow ? 16 : 12.5; tx.FontWeight = isNow ? FontWeights.Bold : FontWeights.SemiBold;
                    tx.Foreground = Br(isNow ? "#1E1A1B" : "#6E6466");
                    Canvas.SetLeft(tx, pts[i].X - 25); Canvas.SetTop(tx, pts[i].Y - (isNow ? 34 : 26)); cv.Children.Add(tx);
                }
            }
            // animación: la línea se dibuja de izquierda a derecha
            if (animateChart)
            {
                animateChart = false;
                Rect r0 = new Rect(0, -50, 0, H + 100), r1 = new Rect(0, -50, W + 30, H + 100);
                RectangleGeometry clip = new RectangleGeometry(r0); cv.Clip = clip;
                RectAnimation ra = new RectAnimation(r0, r1, new Duration(TimeSpan.FromMilliseconds(950)));
                CubicEase ce = new CubicEase(); ce.EasingMode = EasingMode.EaseInOut; ra.EasingFunction = ce;
                clip.BeginAnimation(RectangleGeometry.RectProperty, ra);
            }
            else cv.Clip = null;
        }

        public void UpdateChart()
        {
            string who = StatsWho();
            string mode = U<RadioButton>("ChartDay").IsChecked == true ? "D" : U<RadioButton>("ChartWeek").IsChecked == true ? "W" : "M";
            DateTime today = App.Now().Date;
            CultureInfo es = Ps.ES;
            List<ChartBucket> buckets = new List<ChartBucket>();
            if (mode == "D")
            {
                for (int k = 13; k >= 0; k--) { DateTime d = today.AddDays(-k); buckets.Add(Bucket(d, d.AddDays(1), d.ToString("ddd d", es), d.ToString("dddd d MMMM", es))); }
            }
            else if (mode == "W")
            {
                DateTime mon = today.AddDays(-((((int)today.DayOfWeek) + 6) % 7));
                for (int k = 11; k >= 0; k--)
                {
                    DateTime d = mon.AddDays(-7 * k);
                    int wk = es.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                    buckets.Add(Bucket(d, d.AddDays(7), string.Format(L("chWeekL"), wk), d.ToString("d MMM", es) + " – " + d.AddDays(6).ToString("d MMM", es)));
                }
            }
            else
            {
                DateTime m0 = new DateTime(today.Year, today.Month, 1);
                for (int k = 11; k >= 0; k--) { DateTime d = m0.AddMonths(-k); buckets.Add(Bucket(d, d.AddMonths(1), d.ToString("MMM", es), d.ToString("MMMM yyyy", es))); }
            }
            DateTime start = buckets[0].From, end = buckets[buckets.Count - 1].To;
            if (G.Cache != null)
                foreach (KeyValuePair<string, TramiteData> kv in G.Cache)
                {
                    if (StatsScopeId != null && kv.Key != StatsScopeId) continue;
                    foreach (ClientRow r in kv.Value.Rows)
                    {
                        if (!r.Fecha.HasValue || r.Fecha.Value < start || r.Fecha.Value >= end) continue;
                        if (who != null && !string.Equals(Ps.Str(r.Resp).Trim(), who, StringComparison.CurrentCultureIgnoreCase)) continue;
                        foreach (ChartBucket b in buckets)
                            if (r.Fecha.Value >= b.From && r.Fecha.Value < b.To) { b.N++; b.Bill += r.PriceV; b.Paid += r.PaidV; break; }
                    }
                }
            int sumN = 0; double sumBill = 0;
            foreach (ChartBucket b in buckets) { sumN += b.N; sumBill += b.Bill; }
            ChartBuckets = buckets;
            // etiquetas: periodo y dinero
            Grid lg = U<Grid>("ChartLabels");
            lg.Children.Clear(); lg.ColumnDefinitions.Clear();
            for (int i = 0; i < buckets.Count; i++)
            {
                ChartBucket b = buckets[i];
                ColumnDefinition cd = new ColumnDefinition(); cd.Width = new GridLength(1, GridUnitType.Star); lg.ColumnDefinitions.Add(cd);
                bool isNow = i == buckets.Count - 1;
                StackPanel foot = new StackPanel(); foot.HorizontalAlignment = HorizontalAlignment.Center;
                TextBlock lab = new TextBlock(); lab.Text = b.Label; lab.FontSize = 12.5; lab.HorizontalAlignment = HorizontalAlignment.Center; lab.FontWeight = FontWeights.SemiBold;
                lab.Foreground = Br(isNow ? "#9E1820" : "#3B3536");
                TextBlock mo = new TextBlock(); mo.Text = b.Bill > 0 ? Euro0(b.Bill) : "–"; mo.FontSize = 11.5; mo.HorizontalAlignment = HorizontalAlignment.Center;
                mo.Foreground = Br("#8A7F81"); mo.Margin = new Thickness(0, 2, 0, 0);
                foot.Children.Add(lab); foot.Children.Add(mo);
                Grid.SetColumn(foot, i); lg.Children.Add(foot);
            }
            // comparación con el periodo anterior
            int nb = buckets.Count, diff = buckets[nb - 1].N - buckets[nb - 2].N;
            string sign = diff > 0 ? "▲ +" + diff : diff < 0 ? "▼ " + diff : "= 0";
            TextBlock delta = U<TextBlock>("ChartDelta");
            delta.Text = string.Format(L("chDeltaTxt"), sign, L("chPrev" + mode));
            delta.Foreground = Br(diff > 0 ? "#1B7F45" : diff < 0 ? "#B42318" : "#6E6466");
            Border db = U<Border>("ChartDeltaB");
            db.Background = Br(diff > 0 ? "#E6F4EC" : diff < 0 ? "#FDECEC" : "#F1EDED");
            db.Visibility = Visibility.Visible;
            animateChart = true;
            DrawChartLine();
            U<TextBlock>("ChartTitle").Text = L("chTitle" + mode);
            U<TextBlock>("ChartSub").Text = string.Format(L("chSub" + mode), sumN, Euro0(sumBill));
        }
        static ChartBucket Bucket(DateTime from, DateTime to, string label, string lng) { ChartBucket b = new ChartBucket(); b.From = from; b.To = to; b.Label = label; b.Long = lng; return b; }

        // las cifras suben desde 0 hasta su valor (0,7 s); en las pruebas, directamente el valor
        void StartCountUp(TextBlock[] tbs, double[] to)
        {
            if (countTimer != null) countTimer.Stop();
            if (App.Quiet) { for (int i = 0; i < tbs.Length; i++) tbs[i].Text = Euro0(to[i]); return; }
            DateTime t0 = DateTime.Now;
            foreach (TextBlock tb in tbs) tb.Text = Euro0(0);
            DispatcherTimer tm = new DispatcherTimer(); tm.Interval = TimeSpan.FromMilliseconds(16);
            tm.Tick += delegate
            {
                double p = (DateTime.Now - t0).TotalMilliseconds / 700; if (p >= 1) p = 1;
                double e = 1 - Math.Pow(1 - p, 3);
                for (int i = 0; i < tbs.Length; i++) tbs[i].Text = Euro0(to[i] * e);
                if (p >= 1) tm.Stop();
            };
            countTimer = tm; tm.Start();
        }

        class StatLine { public string Title, Color, Id; public StatAgg A; }
        static void AddAgg(StatAgg t, StatAgg a) { t.N += a.N; t.Bill += a.Bill; t.Paid += a.Paid; t.Pend += a.Pend; t.Net += a.Net; t.Debt += a.Debt; }

        public void UpdateStats()
        {
            if (statsFilling) return;
            Grid g = U<Grid>("StatsGrid");
            g.Children.Clear(); g.RowDefinitions.Clear(); g.ColumnDefinitions.Clear();
            double[] ws = { 2.2, 70, 1, 1, 1, 1, 70, 1.4 };
            for (int i = 0; i < ws.Length; i++) { ColumnDefinition cd = new ColumnDefinition(); cd.Width = new GridLength(ws[i], i == 1 || i == 6 ? GridUnitType.Pixel : GridUnitType.Star); g.ColumnDefinitions.Add(cd); }
            string who = StatsWho();
            DateTime today = App.Now().Date; DateTime? from = null, to = null;
            switch (U<ComboBox>("StatsPeriod").SelectedIndex)
            {
                case 1: from = new DateTime(today.Year, today.Month, 1); to = from.Value.AddMonths(1); break;
                case 2: to = new DateTime(today.Year, today.Month, 1); from = to.Value.AddMonths(-1); break;
                case 3: from = new DateTime(today.Year, 1, 1); to = from.Value.AddYears(1); break;
            }
            string scope = StatsScopeId, scopeTitle = "";
            List<StatLine> lines = new List<StatLine>();
            StatAgg tot = new StatAgg();
            int nTram = 0;
            foreach (Tramite t in Tramite.All)
            {
                if (G.Cache == null || !G.Cache.ContainsKey(t.Id)) continue;
                if (scope != null && t.Id != scope) continue;
                if (t.NoStats && t.Id != scope) continue;    // las tasas son gastos: no suman como facturado
                nTram++;
                string title = t.TitleText();
                if (scope != null) scopeTitle = title;
                OrderedDictionary groups = new OrderedDictionary(StringComparer.CurrentCultureIgnoreCase);    // [ordered]@{}
                foreach (ClientRow r in G.Cache[t.Id].Rows)
                {
                    if (who != null && !string.Equals(Ps.Str(r.Resp).Trim(), who, StringComparison.CurrentCultureIgnoreCase)) continue;
                    if (from.HasValue && (!r.Fecha.HasValue || r.Fecha.Value < from.Value || r.Fecha.Value >= to.Value)) continue;
                    string key = scope != null ? (Ps.Str(r.Status).Trim().Length > 0 ? Ps.Str(r.Status).Trim() : L("stNoStatus")) : t.Id;
                    if (!groups.Contains(key)) groups[key] = new StatAgg();
                    StatAgg a = (StatAgg)groups[key];
                    a.N++; a.Bill += r.PriceV; a.Paid += r.PaidV; a.Pend += r.Pending; a.Net += r.NetV;
                    if (r.Pending > 0) a.Debt++;
                }
                if (scope != null)
                {
                    // t.StatusColors ya viene en el orden en que PowerShell recorre su @{}; sc solo para buscar sin distinguir mayúsculas
                    Hashtable sc = PsH.New(); foreach (KeyValuePair<string, string> kv in t.StatusColors) sc[kv.Key] = kv.Value;
                    foreach (string st0 in t.StatusColors.Keys) if (!groups.Contains(st0)) groups[st0] = new StatAgg();
                    foreach (DictionaryEntry e in groups)
                    {
                        StatAgg a = (StatAgg)e.Value; AddAgg(tot, a);
                        string st = (string)e.Key;
                        StatLine l = new StatLine(); l.Title = st; l.A = a; l.Color = sc.ContainsKey(st) ? (string)sc[st] : "#C9C1C3"; lines.Add(l);
                    }
                }
                else
                {
                    StatAgg a = groups.Contains(t.Id) ? (StatAgg)groups[t.Id] : new StatAgg();
                    AddAgg(tot, a);
                    StatLine l = new StatLine(); l.Title = title; l.A = a; l.Color = "#C4161C"; l.Id = t.Id; lines.Add(l);
                }
            }
            List<StatLine> sorted = new List<StatLine>(lines);
            Ps.Sort(sorted, delegate(StatLine x, StatLine y) { int c = y.A.Bill.CompareTo(x.A.Bill); return c != 0 ? c : y.A.N.CompareTo(x.A.N); });
            double max = 0; foreach (StatLine l in sorted) if (l.A.Bill > max) max = l.A.Bill;
            // cabecera
            int rowIdx = 0;
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(34) });
            string[] heads = { scope != null ? "stStatus" : "stTramite", "stExp", "stBill", "stPaid", "stPend", "stNet", "stMargin", "stShare" };
            for (int c = 0; c < heads.Length; c++)
            {
                TextBlock h = StatCell(L(heads[c]), c, c == 0 || c == 7 ? TextAlignment.Left : TextAlignment.Right, "#6E6466", true, 12.5);
                Grid.SetRow(h, 0); g.Children.Add(h);
            }
            Border line = new Border(); line.BorderBrush = Br("#E9E4E5"); line.BorderThickness = new Thickness(0, 0, 0, 1);
            Grid.SetColumnSpan(line, 8); g.Children.Add(line);
            foreach (StatLine l in sorted)
            {
                rowIdx++;
                StatAgg a = l.A;
                g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(54) });
                bool muted = a.N == 0;
                Border bg = new Border(); bg.BorderBrush = Br("#F1EDED"); bg.BorderThickness = new Thickness(0, 0, 0, 1);
                Grid.SetRow(bg, rowIdx); Grid.SetColumnSpan(bg, 8); g.Children.Add(bg);
                StackPanel name = new StackPanel(); name.Orientation = Orientation.Horizontal; name.Margin = new Thickness(8, 0, 8, 0); name.VerticalAlignment = VerticalAlignment.Center;
                Ellipse dot = new Ellipse(); dot.Width = 8; dot.Height = 8; dot.Margin = new Thickness(0, 0, 10, 0);
                dot.Fill = Br(muted ? "#D9D3D4" : l.Color);
                if (scope != null) { dot.Width = 10; dot.Height = 10; dot.Stroke = Br("#33000000"); dot.StrokeThickness = 0.6; }
                TextBlock nt = new TextBlock(); nt.Text = l.Title; nt.FontSize = 15;
                nt.Foreground = Br(muted ? "#A69C9E" : "#1E1A1B"); if (!muted) nt.FontWeight = FontWeights.SemiBold;
                name.Children.Add(dot); name.Children.Add(nt);
                Grid.SetRow(name, rowIdx); g.Children.Add(name);
                string fg = muted ? "#B8B0B2" : "#2B2526";
                string margin = a.Bill > 0 ? Math.Round(100 * a.Net / a.Bill).ToString(Ps.Inv) + " %" : "–";
                object[][] cells = {
                    new object[] { a.N.ToString(Ps.Inv), 1, fg },
                    new object[] { Euro0(a.Bill), 2, fg },
                    new object[] { Euro0(a.Paid), 3, muted ? fg : "#2E7D32" },
                    new object[] { Euro0(a.Pend), 4, a.Pend > 0 ? "#B42318" : fg },
                    new object[] { Euro0(a.Net), 5, fg },
                    new object[] { margin, 6, fg } };
                foreach (object[] c in cells)
                {
                    int col = (int)c[1];
                    TextBlock tb = StatCell((string)c[0], col, TextAlignment.Right, (string)c[2], !muted && (col == 2 || col == 5), 14.5);
                    Grid.SetRow(tb, rowIdx); g.Children.Add(tb);
                }
                // barra de peso
                Grid bar = new Grid(); bar.Height = 10; bar.Margin = new Thickness(12, 0, 8, 0); bar.VerticalAlignment = VerticalAlignment.Center;
                double pct = max > 0 ? a.Bill / max : 0;
                ColumnDefinition c1 = new ColumnDefinition(); c1.Width = new GridLength(Math.Max(pct, 0.0001), GridUnitType.Star);
                ColumnDefinition c2 = new ColumnDefinition(); c2.Width = new GridLength(Math.Max(1 - pct, 0.0001), GridUnitType.Star);
                bar.ColumnDefinitions.Add(c1); bar.ColumnDefinitions.Add(c2);
                Border track = new Border(); track.Background = Br("#F3EFEF"); track.CornerRadius = new CornerRadius(4);
                Grid.SetColumnSpan(track, 2); bar.Children.Add(track);
                if (pct > 0)
                {
                    Border fill = new Border(); fill.CornerRadius = new CornerRadius(4);
                    LinearGradientBrush lgb = new LinearGradientBrush(); lgb.StartPoint = new Point(0, 0); lgb.EndPoint = new Point(1, 0);
                    lgb.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#D8202A"), 0));
                    lgb.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString("#7E1118"), 1));
                    fill.Background = lgb; bar.Children.Add(fill);
                }
                Grid.SetRow(bar, rowIdx); Grid.SetColumn(bar, 7); g.Children.Add(bar);
                if (l.Id != null)
                {
                    Border hit = new Border();
                    hit.Background = Brushes.Transparent; hit.Cursor = Cursors.Hand; hit.Tag = l.Id;
                    hit.ToolTip = L("stClickHint"); hit.CornerRadius = new CornerRadius(6);
                    hit.MouseEnter += delegate { hit.Background = Br("#0A9E1820"); };
                    hit.MouseLeave += delegate { hit.Background = Brushes.Transparent; };
                    hit.MouseLeftButtonUp += delegate { StatsScopeId = (string)hit.Tag; FillStatsFilters(); UpdateStats(); };
                    Grid.SetRow(hit, rowIdx); Grid.SetColumnSpan(hit, 8); g.Children.Add(hit);
                }
            }
            // total
            rowIdx++;
            g.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) });
            Border tbb = new Border(); tbb.Background = Br("#FAF7F7"); tbb.BorderBrush = Br("#1E1A1B"); tbb.BorderThickness = new Thickness(0, 1.5, 0, 0); tbb.CornerRadius = new CornerRadius(0, 0, 8, 8);
            Grid.SetRow(tbb, rowIdx); Grid.SetColumnSpan(tbb, 8); g.Children.Add(tbb);
            double? tmargin = tot.Bill > 0 ? (double?)Math.Round(100 * tot.Net / tot.Bill) : null;
            object[][] tcells = {
                new object[] { L("stTotal"), 0, "#1E1A1B", TextAlignment.Left }, new object[] { tot.N.ToString(Ps.Inv), 1, "#1E1A1B", TextAlignment.Right },
                new object[] { Euro0(tot.Bill), 2, "#1E1A1B", TextAlignment.Right }, new object[] { Euro0(tot.Paid), 3, "#2E7D32", TextAlignment.Right },
                new object[] { Euro0(tot.Pend), 4, tot.Pend > 0 ? "#B42318" : "#1E1A1B", TextAlignment.Right },
                new object[] { Euro0(tot.Net), 5, "#9E1820", TextAlignment.Right }, new object[] { tmargin.HasValue ? tmargin.Value.ToString(Ps.Inv) + " %" : "–", 6, "#1E1A1B", TextAlignment.Right } };
            foreach (object[] c in tcells)
            {
                TextBlock t2 = StatCell((string)c[0], (int)c[1], (TextAlignment)c[3], (string)c[2], true, 16);
                Grid.SetRow(t2, rowIdx); g.Children.Add(t2);
            }
            // títulos según el alcance
            U<TextBlock>("StatsTable").Text = scope != null ? string.Format(L("statsTableSt"), scopeTitle) : L("statsTable");
            U<TextBlock>("StatsAllTxt").Text = L("stAll");
            U("StatsAllBtn").Visibility = scope != null ? Visibility.Visible : Visibility.Collapsed;
            // tabla: aparece con un fundido suave
            Duration dur = new Duration(TimeSpan.FromMilliseconds(450));
            CubicEase ease = new CubicEase(); ease.EasingMode = EasingMode.EaseOut;
            TranslateTransform tt = new TranslateTransform(0, 12); g.RenderTransform = tt;
            DoubleAnimation fa = new DoubleAnimation(0, 1, dur); fa.EasingFunction = ease;
            DoubleAnimation ya = new DoubleAnimation(12, 0, dur); ya.EasingFunction = ease;
            g.BeginAnimation(UIElement.OpacityProperty, fa);
            tt.BeginAnimation(TranslateTransform.YProperty, ya);
            // tarjetas
            U<TextBlock>("KpiBill").Text = Euro0(tot.Bill);
            U<TextBlock>("KpiBillF").Text = string.Format(L("kpiBillF"), tot.N);
            U<TextBlock>("KpiPaid").Text = Euro0(tot.Paid);
            double pp = tot.Bill > 0 ? Math.Min(1, tot.Paid / tot.Bill) : 0;
            Grid track2 = U<Grid>("KpiPaidTrack");
            track2.ColumnDefinitions[0].Width = new GridLength(Math.Max(pp, 0.0001), GridUnitType.Star);
            track2.ColumnDefinitions[1].Width = new GridLength(Math.Max(1 - pp, 0.0001), GridUnitType.Star);
            U<TextBlock>("KpiPaidF").Text = string.Format(L("kpiPaidF"), Math.Round(100 * pp));
            U<TextBlock>("KpiPend").Text = Euro0(tot.Pend);
            U<TextBlock>("KpiPendF").Text = string.Format(L("kpiPendF"), tot.Debt);
            U<TextBlock>("KpiNet").Text = Euro0(tot.Net);
            U<TextBlock>("KpiNetF").Text = string.Format(L("kpiNetF"), tmargin.HasValue ? tmargin.Value : 0);
            string hhmm = App.Now().ToString("HH:mm");
            U<TextBlock>("StatsSub").Text = scope != null ? string.Format(L("statsSubOne"), scopeTitle, tot.N, hhmm) : string.Format(L("statsSub"), nTram, tot.N, hhmm);
            StartCountUp(new TextBlock[] { U<TextBlock>("KpiBill"), U<TextBlock>("KpiPaid"), U<TextBlock>("KpiPend"), U<TextBlock>("KpiNet") }, new double[] { tot.Bill, tot.Paid, tot.Pend, tot.Net });
            UpdateChart();
        }
    }
}

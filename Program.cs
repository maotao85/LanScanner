using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using ClosedXML.Excel;

namespace LanScanner
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    // Una riga della tabella risultati.
    public class DeviceRow
    {
        public string Ip { get; set; }
        public string Mac { get; set; }
        public string Host { get; set; }
        public string Vendor { get; set; }
        public string Os { get; set; }
        public string User { get; set; }
    }

    // Barra di avanzamento piatta e colorabile (la ProgressBar di Windows non si lascia ricolorare).
    public class FlatProgressBar : Control
    {
        int _min = 0, _max = 100, _val = 0;

        public int Minimum { get => _min; set { _min = value; Invalidate(); } }
        public int Maximum { get => _max; set { _max = value; Invalidate(); } }
        public int Value
        {
            get => _val;
            set { _val = Math.Max(_min, Math.Min(_max, value)); Invalidate(); }
        }
        public Color TrackColor { get; set; } = Color.FromArgb(0x3A, 0x3A, 0x3C);
        public Color FillColor { get; set; } = Color.FromArgb(0x0A, 0x84, 0xFF);

        public FlatProgressBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int radius = Math.Max(1, (Height - 1) / 2);

            var track = new Rectangle(0, 0, Width - 1, Height - 1);
            using (var path = RoundedRect(track, radius))
            using (var b = new SolidBrush(TrackColor))
                g.FillPath(b, path);

            if (_max > _min && _val > _min)
            {
                float frac = (float)(_val - _min) / (_max - _min);
                int w = (int)Math.Round(frac * (Width - 1));
                if (w >= 2)
                {
                    using (var path = RoundedRect(new Rectangle(0, 0, w, Height - 1), radius))
                    using (var b = new SolidBrush(FillColor))
                        g.FillPath(b, path);
                }
            }
        }

        static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
            if (d <= 1) { path.AddRectangle(r); path.CloseFigure(); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    public class MainForm : Form
    {
        // SendARP di iphlpapi.dll: manda una richiesta ARP a destIp e,
        // se l'host risponde, riempie macAddr col suo indirizzo MAC.
        // Non richiede privilegi di amministratore.
        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        static extern int SendARP(int destIp, int srcIp, byte[] macAddr, ref int macAddrLen);

        // --- Controlli UI ---
        ComboBox cmbAdapter;
        TextBox txtBase;
        NumericUpDown numStart, numEnd;
        CheckBox chkDns;
        Button btnScan, btnStop, btnExport, btnExcel;
        FlatProgressBar progress;
        DataGridView grid;
        Label lblStatus;

        readonly BindingList<DeviceRow> results = new BindingList<DeviceRow>();
        CancellationTokenSource cts;
        int lastReportTick;
        const string ByLine = "[by Mauro Tassera]";

        string gatewayIp = "";
        string ouiStatus = "";

        public MainForm()
        {
            BuildUi();
            LoadExternalOui();   // estende i produttori da oui.txt (se presente) e prepara lo stato
            LoadAdapters();      // dopo: così la barra di stato mostra anche il conteggio del DB
        }

        // ---------------------------------------------------------------
        //  Costruzione interfaccia (tutta in codice, niente file designer)
        // ---------------------------------------------------------------
        void BuildUi()
        {
            Text = $"LAN Scanner {ByLine}";
            Width = 1080;
            Height = 580;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9f);
            MinimumSize = new Size(900, 460);
            BackColor = BgDark;
            ForeColor = FgText;
            SetWindowIcon();

            // Header: logo antenna (disegnato) + titolo
            var header = new Panel { Dock = DockStyle.Top, Height = 60, BackColor = BgPanel };
            header.Paint += Header_Paint;
            header.Controls.Add(new Label { Text = "LAN Scanner", Left = 64, Top = 9, Width = 320, Height = 26, ForeColor = FgText, Font = new Font("Segoe UI Semibold", 14f), BackColor = Color.Transparent });
            header.Controls.Add(new Label { Text = $"Scansione IP / MAC della rete locale   —   {ByLine}", Left = 66, Top = 35, Width = 480, Height = 16, ForeColor = FgMuted, Font = new Font("Segoe UI", 8.5f), BackColor = Color.Transparent });

            var top = new Panel { Dock = DockStyle.Top, Height = 96, Padding = new Padding(10), BackColor = BgBar };

            var lblAdapter = new Label { Text = "Scheda di rete:", Left = 10, Top = 14, Width = 90, AutoSize = false, ForeColor = FgText };
            cmbAdapter = new ComboBox { Left = 105, Top = 10, Width = 400, DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = BgInput, ForeColor = FgText };
            cmbAdapter.SelectedIndexChanged += (s, e) => ApplyAdapterSelection();

            var lblBase = new Label { Text = "Base (3 ottetti):", Left = 10, Top = 48, Width = 90, AutoSize = false, ForeColor = FgText };
            txtBase = new TextBox { Left = 105, Top = 44, Width = 120, Text = "192.168.1", BackColor = BgInput, ForeColor = FgText, BorderStyle = BorderStyle.FixedSingle };

            var lblRange = new Label { Text = "Ultimo ottetto da", Left = 240, Top = 48, Width = 100, AutoSize = false, ForeColor = FgText };
            numStart = new NumericUpDown { Left = 342, Top = 44, Width = 60, Minimum = 0, Maximum = 255, Value = 1, BackColor = BgInput, ForeColor = FgText, BorderStyle = BorderStyle.FixedSingle };
            var lblA = new Label { Text = "a", Left = 408, Top = 48, Width = 14, AutoSize = false, ForeColor = FgText };
            numEnd = new NumericUpDown { Left = 424, Top = 44, Width = 60, Minimum = 0, Maximum = 255, Value = 254, BackColor = BgInput, ForeColor = FgText, BorderStyle = BorderStyle.FixedSingle };

            chkDns = new CheckBox { Text = "Dettagli host: OS e utente", Left = 510, Top = 46, Width = 198, Checked = false, ForeColor = FgText, BackColor = Color.Transparent };

            btnScan = new Button { Text = "Scansiona", Left = 710, Top = 10, Width = 80, Height = 28 };
            btnScan.Click += BtnScan_Click;
            btnStop = new Button { Text = "Ferma", Left = 710, Top = 42, Width = 80, Height = 28, Enabled = false };
            btnStop.Click += (s, e) => StopScan();
            StyleButton(btnScan, true);
            StyleButton(btnStop, false);

            top.Controls.AddRange(new Control[]
            {
                lblAdapter, cmbAdapter, lblBase, txtBase, lblRange,
                numStart, lblA, numEnd, chkDns, btnScan, btnStop
            });

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = BgDark
            };
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Indirizzo IP", DataPropertyName = "Ip", FillWeight = 13 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "MAC", DataPropertyName = "Mac", FillWeight = 17 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Nome host", DataPropertyName = "Host", FillWeight = 18 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Produttore", DataPropertyName = "Vendor", FillWeight = 16 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sistema operativo", DataPropertyName = "Os", FillWeight = 16 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Utente loggato", DataPropertyName = "User", FillWeight = 20 });
            grid.DataSource = results;

            // --- Stile della tabella (tema scuro: cambia qui i colori) ---
            grid.EnableHeadersVisualStyles = false;
            grid.BackgroundColor = BgDark;
            grid.ColumnHeadersDefaultCellStyle.BackColor = BgInput;             // intestazione (grigio sobrio)
            grid.ColumnHeadersDefaultCellStyle.ForeColor = FgText;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9f);
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = 32;
            grid.DefaultCellStyle.BackColor = BgPanel;
            grid.DefaultCellStyle.ForeColor = FgText;
            grid.AlternatingRowsDefaultCellStyle.BackColor = AltRow;            // righe alternate
            grid.DefaultCellStyle.SelectionBackColor = Accent;
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.GridColor = GridLine;
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.RowTemplate.Height = 26;

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 56, Padding = new Padding(10), BackColor = BgBar };

            // Barra pulsanti export ancorata a destra: niente posizioni assolute,
            // così i tasti non finiscono mai fuori dalla finestra.
            var btnBar = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 210,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 12, 6, 0),
                BackColor = BgBar
            };
            btnExcel = new Button { Text = "Esporta Excel", Width = 100, Height = 30, Margin = new Padding(4, 0, 4, 0) };
            btnExcel.Click += BtnExcel_Click;
            btnExport = new Button { Text = "Esporta CSV", Width = 92, Height = 30, Margin = new Padding(4, 0, 4, 0) };
            btnExport.Click += BtnExport_Click;
            StyleButton(btnExcel, true);    // tasto d'azione in evidenza
            StyleButton(btnExport, false);
            btnBar.Controls.Add(btnExcel);  // RightToLeft: il primo aggiunto sta più a destra
            btnBar.Controls.Add(btnExport);

            progress = new FlatProgressBar { Left = 10, Top = 10, Width = 520, Height = 14, TrackColor = BgInput, FillColor = AccentBlue };
            lblStatus = new Label { Left = 10, Top = 30, Width = 520, Height = 18, Text = "Pronto.", ForeColor = FgMuted, BackColor = Color.Transparent };

            bottom.Controls.Add(progress);
            bottom.Controls.Add(lblStatus);
            bottom.Controls.Add(btnBar);

            // Footer con copyright
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 24, BackColor = BgDark };
            footer.Controls.Add(new Label
            {
                Text = $"LAN Scanner {ByLine}   —   © 2026 Mauro Tassera",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = FgMuted,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8f)
            });

            // L'ordine di aggiunta determina il docking:
            // header in alto, poi i controlli, la griglia riempie, in fondo barra + footer.
            Controls.Add(grid);
            Controls.Add(bottom);
            Controls.Add(footer);
            Controls.Add(top);
            Controls.Add(header);
        }

        // ---------------------------------------------------------------
        //  Tema e grafica (personalizzabile: cambia qui la palette)
        // ---------------------------------------------------------------
        static readonly Color BgDark    = ColorTranslator.FromHtml("#232325"); // finestra / sfondo griglia / footer
        static readonly Color BgBar     = ColorTranslator.FromHtml("#2A2A2C"); // barre superiore e inferiore
        static readonly Color BgPanel   = ColorTranslator.FromHtml("#323234"); // header / celle tabella
        static readonly Color BgInput   = ColorTranslator.FromHtml("#3A3A3C"); // campi, intestazione tabella, tasti secondari
        static readonly Color FgText    = ColorTranslator.FromHtml("#ECECEE");
        static readonly Color FgMuted   = ColorTranslator.FromHtml("#9A9A9F");
        static readonly Color Accent    = ColorTranslator.FromHtml("#8A8A8F"); // grafite: logo, selezione
        static readonly Color BtnPrimary = ColorTranslator.FromHtml("#C7C7CC"); // tasto d'azione grigio chiaro (stile mac)
        static readonly Color AccentBlue = ColorTranslator.FromHtml("#0A84FF"); // azzurro della barra di avanzamento
        static readonly Color AltRow    = ColorTranslator.FromHtml("#2A2A2C");
        static readonly Color GridLine  = ColorTranslator.FromHtml("#3C3C40");

        void StyleButton(Button b, bool primary)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 1;
            b.Font = new Font("Segoe UI", 9f);
            b.Cursor = Cursors.Hand;
            if (primary)
            {
                b.BackColor = BtnPrimary;                 // grigio chiaro, tasto d'azione
                b.ForeColor = ColorTranslator.FromHtml("#1A1A1A");
                b.FlatAppearance.BorderColor = ControlPaint.Dark(BtnPrimary);
                b.FlatAppearance.MouseOverBackColor = ControlPaint.Light(BtnPrimary);
            }
            else
            {
                b.BackColor = BgInput;
                b.ForeColor = FgText;
                b.FlatAppearance.BorderColor = GridLine;
                b.FlatAppearance.MouseOverBackColor = ControlPaint.Light(BgInput);
            }
        }

        // Disegna il logo "antenna" nell'header.
        void Header_Paint(object sender, PaintEventArgs e)
        {
            DrawAntenna(e.Graphics, new Rectangle(16, 12, 36, 36), Accent);
        }

        // Antenna broadcast (palo, base, onde di segnale). Vettoriale: nessun file immagine.
        static void DrawAntenna(Graphics g, Rectangle box, Color color)
        {
            var saved = g.SmoothingMode;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(color, 2.2f))
            using (var brush = new SolidBrush(color))
            {
                float cx = box.Left + box.Width / 2f;
                float topY = box.Top + box.Height * 0.30f;
                float botY = box.Bottom - 2;

                // Gambe del palo
                g.DrawLine(pen, cx, topY, cx - box.Width * 0.18f, botY);
                g.DrawLine(pen, cx, topY, cx + box.Width * 0.18f, botY);
                // Traversa
                g.DrawLine(pen, cx - box.Width * 0.12f, botY - box.Height * 0.16f,
                                 cx + box.Width * 0.12f, botY - box.Height * 0.16f);
                // Punta
                g.FillEllipse(brush, cx - 3f, topY - 3f, 6f, 6f);
                // Onde di segnale su entrambi i lati
                using (var wave = new Pen(color, 1.8f))
                {
                    var r1 = new RectangleF(cx - 9, topY - 9, 18, 18);
                    var r2 = new RectangleF(cx - 14, topY - 14, 28, 28);
                    g.DrawArc(wave, r1, 200, 50);
                    g.DrawArc(wave, r1, 290, 50);
                    g.DrawArc(wave, r2, 205, 40);
                    g.DrawArc(wave, r2, 295, 40);
                }
            }
            g.SmoothingMode = saved;
        }

        // Icona della finestra/taskbar: stessa antenna, disegnata su bitmap 32x32.
        void SetWindowIcon()
        {
            // Usa l'icona incorporata nell'eseguibile (antenna.ico): titolo e taskbar uguali all'exe.
            try
            {
                var ico = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (ico != null) { this.Icon = ico; return; }
            }
            catch { }
            // Ripiego: la disegna al volo se l'icona incorporata non è disponibile.
            try
            {
                using (var bmp = new Bitmap(32, 32))
                {
                    using (var g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.Transparent);
                        DrawAntenna(g, new Rectangle(2, 2, 28, 28), Accent);
                    }
                    this.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon());
                }
            }
            catch { /* icona non critica */ }
        }

        // ---------------------------------------------------------------
        //  Rilevamento schede di rete attive (IPv4)
        // ---------------------------------------------------------------
        void LoadAdapters()
        {
            cmbAdapter.Items.Clear();
            foreach (var ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (ni.OperationalStatus != OperationalStatus.Up) continue;
                if (ni.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;

                foreach (var ua in ni.GetIPProperties().UnicastAddresses)
                {
                    if (ua.Address.AddressFamily != AddressFamily.InterNetwork) continue;
                    var gw = ni.GetIPProperties().GatewayAddresses
                               .FirstOrDefault(g => g.Address != null && g.Address.AddressFamily == AddressFamily.InterNetwork)
                               ?.Address;
                    cmbAdapter.Items.Add(new AdapterItem
                    {
                        Name = ni.Name,
                        Description = ni.Description,
                        Address = ua.Address,
                        Mask = ua.IPv4Mask,
                        Gateway = gw
                    });
                }
            }
            if (cmbAdapter.Items.Count > 0)
                cmbAdapter.SelectedIndex = 0; // applica la prima -> compila base e range
        }

        class AdapterItem
        {
            public string Name;
            public string Description;
            public IPAddress Address;
            public IPAddress Mask;
            public IPAddress Gateway;
            public override string ToString() => $"{Name} — {Address}  ({Description})";
        }

        // Compila base e range in funzione della scheda selezionata.
        void ApplyAdapterSelection()
        {
            if (cmbAdapter.SelectedItem is AdapterItem a && a.Address != null)
            {
                var b = a.Address.GetAddressBytes();
                txtBase.Text = $"{b[0]}.{b[1]}.{b[2]}";
                numStart.Value = 1;
                numEnd.Value = 254;
                gatewayIp = a.Gateway?.ToString() ?? "";

                int prefix = MaskToPrefix(a.Mask);
                string sub = (prefix == 24)
                    ? $"Subnet /24 rilevata ({a.Address}/{a.Mask})."
                    : $"Nota: subnet /{prefix} (mask {a.Mask}) — il range /24 potrebbe non coprire tutta la rete.";
                lblStatus.Text = string.IsNullOrEmpty(ouiStatus) ? sub : sub + "   •   " + ouiStatus;
            }
        }

        static int MaskToPrefix(IPAddress mask)
        {
            if (mask == null) return 0;
            int count = 0;
            foreach (var by in mask.GetAddressBytes())
            {
                byte v = by;
                while (v != 0) { count += v & 1; v >>= 1; }
            }
            return count;
        }

        // Interrompe la scansione con feedback immediato (un solo clic).
        void StopScan()
        {
            if (cts == null || !btnStop.Enabled) return;
            btnStop.Enabled = false;                  // disabilita subito: niente clic ripetuti
            lblStatus.Text = "Interruzione in corso…";
            try { cts.Cancel(); } catch { }
        }

        // ---------------------------------------------------------------
        //  Scansione
        // ---------------------------------------------------------------
        async void BtnScan_Click(object sender, EventArgs e)
        {
            List<IPAddress> ips;
            try
            {
                ips = BuildIpList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Input non valido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (ips.Count == 0)
            {
                MessageBox.Show("Il range selezionato non contiene indirizzi.", "Nulla da scansionare",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            results.Clear();
            cts = new CancellationTokenSource();
            var token = cts.Token;
            bool doDns = chkDns.Checked;
            string gw = gatewayIp;

            progress.Minimum = 0;
            progress.Maximum = ips.Count;
            progress.Value = 0;
            int done = 0;
            lastReportTick = 0;

            btnScan.Enabled = false;
            btnStop.Enabled = true;
            cmbAdapter.Enabled = false;

            try
            {
                await Task.Run(() =>
                {
                    var po = new ParallelOptions { MaxDegreeOfParallelism = 64 };
                    Parallel.ForEach(ips, po, ip =>
                    {
                        if (token.IsCancellationRequested) return;   // stop pulito, senza eccezioni

                        string mac = TryGetMac(ip);
                        if (mac != null)
                        {
                            string vendor = OuiVendor(mac);
                            string os = ClassifyOs(ip, mac, vendor, gw, doDns);
                            string host = "", user = "";
                            if (doDns)
                            {
                                host = TryGetHost(ip);
                                // Utente loggato solo dove c'è un vero nome host (computer come SIEMENS-GE),
                                // non su stampanti/apparati senza nome. Richiede diritti admin sul target.
                                bool named = !string.IsNullOrEmpty(host) && host != "Sconosciuto";
                                if (named) user = TryGetLoggedUser(host);
                            }
                            var row = new DeviceRow { Ip = ip.ToString(), Mac = mac, Host = host, Vendor = vendor, Os = os, User = user };
                            TryBeginInvoke(() => results.Add(row));
                        }

                        int d = Interlocked.Increment(ref done);
                        // Aggiorna la UI al massimo ~16 volte/sec: evita di intasare il thread
                        // dell'interfaccia e tiene il tasto Ferma sempre reattivo.
                        int nowTick = Environment.TickCount;
                        if (d == ips.Count || nowTick - Volatile.Read(ref lastReportTick) >= 60)
                        {
                            Volatile.Write(ref lastReportTick, nowTick);
                            TryBeginInvoke(() =>
                            {
                                progress.Value = Math.Min(d, progress.Maximum);
                                lblStatus.Text = $"Scansione... {d}/{ips.Count} — trovati {results.Count}";
                            });
                        }
                    });
                });

                SortByIp();
                lblStatus.Text = token.IsCancellationRequested
                    ? $"Interrotto — {results.Count} dispositivi trovati finora."
                    : $"Completato — {results.Count} dispositivi trovati su {ips.Count} indirizzi.";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Errore: " + ex.Message;
            }
            finally
            {
                btnScan.Enabled = true;
                btnStop.Enabled = false;
                cmbAdapter.Enabled = true;
            }
        }

        // Costruisce la lista di IP dal base "a.b.c" + range dell'ultimo ottetto.
        List<IPAddress> BuildIpList()
        {
            string baseText = (txtBase.Text ?? "").Trim();
            var parts = baseText.Split('.');
            if (parts.Length != 3)
                throw new Exception("La base deve essere formata da 3 ottetti, es. 192.168.1");

            var octets = new int[3];
            for (int i = 0; i < 3; i++)
            {
                if (!int.TryParse(parts[i], out octets[i]) || octets[i] < 0 || octets[i] > 255)
                    throw new Exception($"Ottetto non valido: \"{parts[i]}\" (ammessi 0-255).");
            }

            int start = (int)numStart.Value;
            int end = (int)numEnd.Value;
            if (start > end) { int t = start; start = end; end = t; }

            var list = new List<IPAddress>(end - start + 1);
            for (int last = start; last <= end; last++)
                list.Add(new IPAddress(new byte[] { (byte)octets[0], (byte)octets[1], (byte)octets[2], (byte)last }));
            return list;
        }

        // Richiesta ARP: ritorna il MAC formattato, o null se l'host non risponde.
        static string TryGetMac(IPAddress ip)
        {
            try
            {
                byte[] mac = new byte[6];
                int len = mac.Length;
                int dest = BitConverter.ToInt32(ip.GetAddressBytes(), 0); // network byte order su little-endian
                int res = SendARP(dest, 0, mac, ref len);
                if (res == 0 && len >= 6)
                    return string.Join(":", mac.Take(6).Select(b => b.ToString("X2")));
            }
            catch { /* ignora: host non raggiungibile */ }
            return null;
        }

        // Reverse DNS (può essere lento o fallire: opzionale).
        static string TryGetHost(IPAddress ip)
        {
            try
            {
                string h = Dns.GetHostEntry(ip).HostName;
                return string.IsNullOrWhiteSpace(h) ? "Sconosciuto" : h;
            }
            catch { return "Sconosciuto"; }
        }

        // Classifica tipo/OS combinando: gateway, produttore (OUI), MAC casuale, poi TTL + porte.
        static string ClassifyOs(IPAddress ip, string mac, string vendor, string gatewayIp, bool deep)
        {
            // 1) È il router? Lo so con certezza: è il gateway della scheda di rete.
            if (!string.IsNullOrEmpty(gatewayIp) && ip.ToString() == gatewayIp)
                return "Router / Gateway";

            // 2) Indizio dal produttore (MAC)
            string v = (vendor ?? "").ToLowerInvariant();
            if (v.Contains("apple")) return "Apple (iOS/macOS)";
            if (v.Contains("cisco") || v.Contains("mikrotik") || v.Contains("ubiquiti") || v.Contains("tp-link")
                || v.Contains("netgear") || v.Contains("avm") || v.Contains("zyxel") || v.Contains("d-link")
                || v.Contains("aruba") || v.Contains("juniper") || v.Contains("fortinet") || v.Contains("huawei"))
                return "Apparato di rete";
            if (v.Contains("brother") || v.Contains("canon") || v.Contains("epson") || v.Contains("kyocera"))
                return "Stampante";
            if (v.Contains("espressif") || v.Contains("raspberry") || v.Contains("tuya"))
                return "Dispositivo IoT";
            if (v.Contains("vmware") || v.Contains("virtualbox") || v.Contains("qemu") || v.Contains("parallels"))
                return "Macchina virtuale";

            // 3) Fingerprinting: TTL (+ porte se analisi approfondita)
            int ttl = PingTtl(ip);
            if (deep)
            {
                if (PortOpen(ip, 3389, 300)) return "Windows";            // RDP
                bool smb = PortOpen(ip, 445, 300);
                bool ssh = PortOpen(ip, 22, 300);
                if (smb) return (ttl > 0 && ttl <= 64 && ssh) ? "Linux/Unix (Samba)" : "Windows";
                if (ssh) return "Linux/Unix";
            }
            if (ttl > 0)
            {
                if (ttl <= 64) return IsRandomMac(mac) ? "Mobile/altro (MAC privato)" : "Linux/Unix/macOS";
                if (ttl <= 128) return "Windows";
                return "Apparato di rete";
            }
            return IsRandomMac(mac) ? "Mobile/altro (MAC privato)" : "Sconosciuto";
        }

        // --- Produttore dal MAC (OUI) ---
        // true se il MAC è "locally administered" (bit 0x02 del primo ottetto):
        // tipico dei MAC casuali/privacy dei telefoni moderni (iPhone, Android).
        static bool IsRandomMac(string mac)
        {
            try
            {
                int b = Convert.ToInt32(mac.Split(':', '-')[0], 16);
                return (b & 0x02) != 0;
            }
            catch { return false; }
        }

        static string OuiVendor(string mac)
        {
            if (string.IsNullOrEmpty(mac)) return "";
            try
            {
                var p = mac.Replace(":", "").Replace("-", "").ToUpperInvariant();
                if (p.Length >= 6 && OuiMap.TryGetValue(p.Substring(0, 6), out var v)) return v;
                if (IsRandomMac(mac)) return "MAC casuale (privacy)";
            }
            catch { }
            return "";
        }

        // Estende la tabella produttori da un file "oui.txt" accanto all'exe (facoltativo).
        // Accetta sia il formato semplice "ACBC32 Apple",
        // sia il file "manuf" di Wireshark così com'è: "AC:BC:32<TAB>Apple<TAB>Apple, Inc."
        // (gestisce anche i prefissi con maschera tipo "AC:BC:32:00:00/36").
        void LoadExternalOui()
        {
            bool fileFound = false;
            try
            {
                var path = Path.Combine(AppContext.BaseDirectory, "oui.txt");
                if (File.Exists(path))
                {
                    fileFound = true;
                    foreach (var raw in File.ReadLines(path))
                    {
                        var s = raw.Trim();
                        if (s.Length < 8 || s.StartsWith("#")) continue;

                        var parts = s.Split('\t');                 // formato Wireshark manuf (tab)
                        if (parts.Length < 2)
                            parts = s.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length < 2) continue;

                        var prefix = parts[0];
                        int slash = prefix.IndexOf('/');           // ignora /28 /36 (MA-M / MA-S)
                        if (slash >= 0) prefix = prefix.Substring(0, slash);
                        var key = prefix.Replace(":", "").Replace("-", "").ToUpperInvariant();
                        if (key.Length >= 6) OuiMap[key.Substring(0, 6)] = parts[1].Trim();
                    }
                }
            }
            catch { }
            ouiStatus = fileFound
                ? $"DB produttori: {OuiMap.Count} voci (oui.txt caricato)"
                : $"DB produttori: {OuiMap.Count} voci (integrato; oui.txt non trovato)";
        }

        static readonly Dictionary<string, string> OuiMap = BuildOui();
        static Dictionary<string, string> BuildOui()
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in OuiData.Split('\n'))
            {
                var s = line.Trim();
                int sp = s.IndexOf(' ');
                if (sp < 6) continue;
                var key = s.Substring(0, sp).Replace(":", "").ToUpperInvariant();
                if (key.Length == 6) d[key] = s.Substring(sp + 1).Trim();
            }
            return d;
        }

        // Tabella OUI integrata con i produttori più comuni (estendibile con oui.txt).
        const string OuiData = @"
AC:BC:32 Apple
F0:DB:F8 Apple
F0:DC:E2 Apple
F4:0F:24 Apple
DC:A9:04 Apple
88:66:5A Apple
90:B2:1F Apple
A4:D1:D2 Apple
3C:07:54 Apple
00:1E:C2 Apple
00:25:00 Apple
00:26:08 Apple
00:26:BB Apple
00:1C:B3 Apple
00:1B:63 Apple
00:23:DF Apple
00:24:36 Apple
E4:25:E7 Apple
7C:D1:C3 Apple
B8:C7:5D Apple
8C:7B:9D Apple
28:CF:DA Apple
28:E0:2C Apple
D0:03:4B Apple
C8:2A:14 Apple
BC:92:6B Apple
60:FA:CD Apple
A8:60:B6 Apple
34:15:9E Apple
40:A6:D9 Apple
58:55:CA Apple
64:20:0C Apple
64:76:BA Apple
A4:B1:97 Apple
A4:C3:61 Apple
AC:29:3A Apple
B0:65:BD Apple
B4:18:D1 Apple
D8:96:95 Apple
DC:2B:2A Apple
E0:F8:47 Apple
F8:1E:DF Apple
FC:25:3F Apple
F4:1B:A1 Apple
8C:77:12 Samsung
38:AA:3C Samsung
5C:0A:5B Samsung
1C:62:B8 Samsung
34:23:BA Samsung
D0:17:6A Samsung
78:1F:DB Samsung
08:08:C2 Samsung
00:12:FB Samsung
00:21:19 Samsung
F0:B4:29 Xiaomi
28:6C:07 Xiaomi
34:CE:00 Xiaomi
50:8F:4C Xiaomi
64:09:80 Xiaomi
78:11:DC Xiaomi
8C:BE:BE Xiaomi
64:CC:2E Xiaomi
00:E0:FC Huawei
28:6E:D4 Huawei
48:46:FB Huawei
80:FB:06 Huawei
24:09:95 Huawei
00:1B:21 Intel
00:1E:67 Intel
3C:A9:F4 Intel
7C:7A:91 Intel
94:65:9C Intel
A0:A8:CD Intel
AC:72:89 Intel
00:15:17 Intel
00:14:22 Dell
18:66:DA Dell
B8:CA:3A Dell
F8:BC:12 Dell
50:9A:4C Dell
3C:4A:92 Hewlett Packard
70:10:6F Hewlett Packard
A0:48:1C Hewlett Packard
94:57:A5 Hewlett Packard
30:8D:99 Hewlett Packard
80:E8:2C Hewlett Packard
2C:44:FD Hewlett Packard
50:C7:BF TP-Link
A4:2B:B0 TP-Link
AC:84:C6 TP-Link
EC:08:6B TP-Link
F4:F2:6D TP-Link
C0:4A:00 TP-Link
14:CC:20 TP-Link
00:1D:0F TP-Link
00:09:5B Netgear
00:14:6C Netgear
20:4E:7F Netgear
2C:B0:5D Netgear
A0:63:91 Netgear
C4:04:15 Netgear
44:94:FC Netgear
00:1F:3F AVM
08:96:D7 AVM
C0:25:06 AVM
E0:28:6D AVM
00:0C:42 MikroTik
4C:5E:0C MikroTik
6C:3B:6B MikroTik
B8:69:F4 MikroTik
CC:2D:E0 MikroTik
DC:2C:6E MikroTik
E4:8D:8C MikroTik
48:A9:8A MikroTik
00:15:6D Ubiquiti
04:18:D6 Ubiquiti
24:A4:3C Ubiquiti
44:D9:E7 Ubiquiti
78:8A:20 Ubiquiti
80:2A:A8 Ubiquiti
F0:9F:C2 Ubiquiti
FC:EC:DA Ubiquiti
68:72:51 Ubiquiti
74:AC:B9 Ubiquiti
00:00:0C Cisco
00:1A:A1 Cisco
00:25:45 Cisco
F4:CF:E2 Cisco
00:1B:11 D-Link
1C:7E:E5 D-Link
28:10:7B D-Link
C8:BE:19 D-Link
00:80:77 Brother
30:05:5C Brother
00:1E:8F Canon
2C:9E:FC Canon
00:00:48 Epson
64:EB:8C Epson
A4:EE:57 Epson
B8:27:EB Raspberry Pi
DC:A6:32 Raspberry Pi
E4:5F:01 Raspberry Pi
28:CD:C1 Raspberry Pi
24:0A:C4 Espressif
30:AE:A4 Espressif
3C:71:BF Espressif
84:0D:8E Espressif
A0:20:A6 Espressif
B4:E6:2D Espressif
7C:DF:A1 Espressif
24:62:AB Espressif
C4:4F:33 Espressif
52:54:00 QEMU/KVM
00:0C:29 VMware
00:50:56 VMware
00:05:69 VMware
00:1C:14 VMware
08:00:27 VirtualBox
";

        // Legge il TTL dalla risposta ping (Windows ~128, Linux/macOS ~64, apparati ~255).
        static int PingTtl(IPAddress ip)
        {
            try
            {
                using (var p = new Ping())
                {
                    var r = p.Send(ip, 600);
                    if (r != null && r.Status == IPStatus.Success && r.Options != null)
                        return r.Options.Ttl;
                }
            }
            catch { }
            return 0;
        }

        // Prova una connessione TCP con timeout: true se la porta è aperta.
        static bool PortOpen(IPAddress ip, int port, int timeoutMs)
        {
            try
            {
                using (var c = new TcpClient())
                {
                    var ar = c.BeginConnect(ip, port, null, null);
                    if (!ar.AsyncWaitHandle.WaitOne(timeoutMs)) return false;
                    c.EndConnect(ar);
                    return true;
                }
            }
            catch { return false; }
        }

        // Utente attualmente loggato su una macchina Windows remota, via WMI.
        // Usa le credenziali correnti: funziona dove hai diritti admin/di dominio sul target.
        static string TryGetLoggedUser(string target)
        {
            return Bounded(() =>
            {
                try
                {
                    var options = new ConnectionOptions
                    {
                        Timeout = TimeSpan.FromSeconds(4),
                        Impersonation = ImpersonationLevel.Impersonate
                    };
                    var scope = new ManagementScope($"\\\\{target}\\root\\cimv2", options);
                    scope.Connect();
                    var q = new ObjectQuery("SELECT UserName FROM Win32_ComputerSystem");
                    using (var s = new ManagementObjectSearcher(scope, q))
                    {
                        foreach (ManagementObject mo in s.Get())
                        {
                            var u = mo["UserName"] as string;
                            return string.IsNullOrWhiteSpace(u) ? "—" : u;
                        }
                    }
                }
                catch { /* firewall/DCOM o niente permessi */ }
                return "—";
            }, 6000);
        }

        // Esegue una funzione con un tetto di tempo, così una chiamata WMI bloccata non congela la riga.
        static string Bounded(Func<string> f, int ms)
        {
            try
            {
                var t = Task.Run(f);
                return t.Wait(ms) ? t.Result : "—";
            }
            catch { return "—"; }
        }

        void SortByIp()
        {
            var sorted = results.OrderBy(r => IpToKey(r.Ip)).ToList();
            results.Clear();
            foreach (var r in sorted) results.Add(r);
        }

        static uint IpToKey(string ip)
        {
            var b = IPAddress.Parse(ip).GetAddressBytes();
            return ((uint)b[0] << 24) | ((uint)b[1] << 16) | ((uint)b[2] << 8) | b[3];
        }

        // Invoke sicuro (ignora se il form sta chiudendo).
        void TryBeginInvoke(Action a)
        {
            try { if (IsHandleCreated) BeginInvoke(a); }
            catch { }
        }

        // ---------------------------------------------------------------
        //  Esportazione CSV
        // ---------------------------------------------------------------
        void BtnExport_Click(object sender, EventArgs e)
        {
            if (results.Count == 0)
            {
                MessageBox.Show("Nessun risultato da esportare.", "Esporta CSV",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (var sfd = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "lan-scan.csv" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                var sb = new StringBuilder();
                sb.AppendLine("IP,MAC,Host,Produttore,Sistema operativo,Utente loggato");
                foreach (var r in results)
                    sb.AppendLine($"{r.Ip},{r.Mac},{Csv(r.Host)},{Csv(r.Vendor)},{Csv(r.Os)},{Csv(r.User)}");
                sb.AppendLine();
                sb.AppendLine($"Generato con LAN Scanner {ByLine}");
                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                lblStatus.Text = "Esportato in " + sfd.FileName;
            }
        }

        // ---------------------------------------------------------------
        //  Esportazione Excel (.xlsx) con ClosedXML
        // ---------------------------------------------------------------
        void BtnExcel_Click(object sender, EventArgs e)
        {
            if (results.Count == 0)
            {
                MessageBox.Show("Nessun risultato da esportare.", "Esporta Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (var sfd = new SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx", FileName = "lan-scan.xlsx" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                try
                {
                    ExportExcel(sfd.FileName);
                    lblStatus.Text = "Esportato in " + sfd.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore durante l'esportazione: " + ex.Message, "Esporta Excel",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        void ExportExcel(string path)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Scansione LAN");

                // Titolo
                ws.Cell(1, 1).Value = $"Scansione LAN  {ByLine}";
                ws.Range(1, 1, 1, 6).Merge();
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Font.FontSize = 14;
                ws.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#1F6FEB");

                // Metadati della scansione
                ws.Cell(2, 1).Value = "Data";
                ws.Cell(2, 2).Value = DateTime.Now;
                ws.Cell(2, 2).Style.DateFormat.Format = "yyyy-mm-dd hh:mm";
                ws.Cell(3, 1).Value = "Range";
                ws.Cell(3, 2).Value = $"{txtBase.Text}.{(int)numStart.Value}-{(int)numEnd.Value}";
                ws.Cell(4, 1).Value = "Dispositivi";
                ws.Cell(4, 2).Value = results.Count;
                ws.Range(2, 1, 4, 1).Style.Font.Bold = true;

                // Intestazione tabella
                const int headerRow = 6;
                ws.Cell(headerRow, 1).Value = "Indirizzo IP";
                ws.Cell(headerRow, 2).Value = "MAC";
                ws.Cell(headerRow, 3).Value = "Nome host";
                ws.Cell(headerRow, 4).Value = "Produttore";
                ws.Cell(headerRow, 5).Value = "Sistema operativo";
                ws.Cell(headerRow, 6).Value = "Utente loggato";

                // Righe dati
                int r = headerRow + 1;
                foreach (var d in results)
                {
                    ws.Cell(r, 1).Value = d.Ip;
                    ws.Cell(r, 2).Value = d.Mac;
                    ws.Cell(r, 3).Value = d.Host ?? "";
                    ws.Cell(r, 4).Value = d.Vendor ?? "";
                    ws.Cell(r, 5).Value = d.Os ?? "";
                    ws.Cell(r, 6).Value = d.User ?? "";
                    r++;
                }

                // Trasforma in tabella Excel (stile + autofiltro) e congela l'intestazione
                var tableRange = ws.Range(headerRow, 1, r - 1, 6);
                var table = tableRange.CreateTable();
                table.Theme = XLTableTheme.TableStyleMedium2;

                ws.Columns().AdjustToContents();
                ws.SheetView.FreezeRows(headerRow);

                wb.SaveAs(path);
            }
        }

        static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\""))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}

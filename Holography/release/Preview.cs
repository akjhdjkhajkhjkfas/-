using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace HolographyPreview
{
    public static class Codebook
    {
        public const int Rows = 40, Columns = 64;

        public static bool[,] Validate(IList<IList<string>> values)
        {
            if (values.Count != Rows || values.Any(row => row.Count != Columns))
                throw new FormatException("需要 40 行 × 64 列的码本；当前为 " + values.Count +
                    " 行，首行 " + (values.Count == 0 ? 0 : values[0].Count) +
                    " 列。不会自动裁剪、补零或重排原始数据。");
            var result = new bool[Rows, Columns];
            for (int r = 0; r < Rows; r++)
                for (int c = 0; c < Columns; c++)
                {
                    double value;
                    if (!double.TryParse(values[r][c].Trim(), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out value) || (value != 0 && value != 1))
                        throw new FormatException(string.Format("第 {0} 行、第 {1} 列不是 0 或 1：{2}",
                            r + 1, c + 1, values[r][c]));
                    result[r, c] = value == 1;
                }
            return result;
        }

        public static bool[,] ReadText(string path)
        {
            string text;
            using (var reader = new StreamReader(path, Encoding.UTF8, true)) text = reader.ReadToEnd();
            var rows = new List<IList<string>>();
            foreach (var line in text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n').Split('\n'))
            {
                char delimiter = line.Contains('\t') ? '\t' : ',';
                rows.Add(line.Split(delimiter).Select(x => x.Trim().Trim('"')).ToArray());
            }
            return Validate(rows);
        }

        // Matches the original Excel::ExcelRead approach: first worksheet, UsedRange, Value.
        // All COM objects belong to this worker and are released, including on failure.
        public static bool[,] ReadExcel(string path)
        {
            object app = null, books = null, book = null, sheets = null, sheet = null, range = null;
            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null) throw new InvalidOperationException("读取 Excel 文件需要安装 Microsoft Excel，也可另存为 CSV 后导入。");
                app = Activator.CreateInstance(type);
                dynamic excel = app;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.AutomationSecurity = 3;
                books = excel.Workbooks;
                book = ((dynamic)books).Open(Path.GetFullPath(path), 0, true);
                sheets = ((dynamic)book).Worksheets;
                sheet = ((dynamic)sheets).Item[1];
                range = ((dynamic)sheet).UsedRange;
                var cells = ((dynamic)range).Value2 as Array;
                if (cells == null || cells.Rank != 2) throw new FormatException("首个工作表没有二维码本数据。");
                var rows = new List<IList<string>>();
                for (int r = cells.GetLowerBound(0); r <= cells.GetUpperBound(0); r++)
                {
                    var row = new List<string>();
                    for (int c = cells.GetLowerBound(1); c <= cells.GetUpperBound(1); c++)
                        row.Add(Convert.ToString(cells.GetValue(r, c), CultureInfo.InvariantCulture));
                    rows.Add(row);
                }
                return Validate(rows);
            }
            finally
            {
                if (book != null) { try { ((dynamic)book).Close(false); } catch { } }
                if (app != null) { try { ((dynamic)app).Quit(); } catch { } }
                foreach (var obj in new[] { range, sheet, sheets, book, books, app })
                    if (obj != null && Marshal.IsComObject(obj)) Marshal.FinalReleaseComObject(obj);
            }
        }

        public static bool[,] Read(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".xls" || ext == ".xlsx") return ReadExcel(path);
            if (ext == ".csv" || ext == ".tsv" || ext == ".txt") return ReadText(path);
            throw new FormatException("支持 .xls、.xlsx、.csv、.tsv 和 .txt 文件。");
        }

        public static void Save(string path, bool[,] data)
        {
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
                for (int r = 0; r < Rows; r++)
                    writer.WriteLine(string.Join(",", Enumerable.Range(0, Columns)
                        .Select(c => data[r, c] ? "1" : "0")));
        }

        public static bool[,] Demo()
        {
            var data = new bool[Rows, Columns];
            for (int r = 0; r < Rows; r++)
                for (int c = 0; c < Columns; c++) data[r, c] = (r / 4 + c / 8) % 2 == 0;
            return data;
        }
    }

    public sealed class MatrixView : Control
    {
        public bool[,] Data = new bool[Codebook.Rows, Codebook.Columns];
        public event Action Changed;
        public event Action<int, int> HoverCell;
        public bool EditingEnabled;
        public MatrixView()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            MinimumSize = new Size(360, 270);
        }

        public Rectangle GridBounds
        {
            get
            {
                int cell = Math.Max(1, Math.Min((ClientSize.Width - 54) / Codebook.Columns,
                    (ClientSize.Height - 48) / Codebook.Rows));
                int w = cell * Codebook.Columns, h = cell * Codebook.Rows;
                return new Rectangle(38 + (ClientSize.Width - 48 - w) / 2,
                    28 + (ClientSize.Height - 40 - h) / 2, w, h);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var rect = GridBounds;
            int size = rect.Width / Codebook.Columns;
            using (var on = new SolidBrush(Color.FromArgb(255, 111, 35)))
            using (var line = new Pen(Color.FromArgb(112, 120, 128)))
            using (var text = new SolidBrush(Color.FromArgb(83, 96, 111)))
            using (var format = new StringFormat { Alignment = StringAlignment.Center })
            {
                for (int r = 0; r < Codebook.Rows; r++)
                    for (int c = 0; c < Codebook.Columns; c++)
                        if (Data[r, c]) e.Graphics.FillRectangle(on, rect.X + c * size, rect.Y + r * size, size, size);
                for (int r = 0; r <= Codebook.Rows; r++)
                    e.Graphics.DrawLine(line, rect.Left, rect.Top + r * size, rect.Right, rect.Top + r * size);
                for (int c = 0; c <= Codebook.Columns; c++)
                    e.Graphics.DrawLine(line, rect.Left + c * size, rect.Top, rect.Left + c * size, rect.Bottom);
                foreach (int c in new[] { 1, 8, 16, 24, 32, 40, 48, 56, 64 })
                    e.Graphics.DrawString(c.ToString(), Font, text,
                        rect.Left + (c - 0.5f) * size, rect.Top - 22, format);
                foreach (int r in new[] { 1, 8, 16, 24, 32, 40 })
                    e.Graphics.DrawString(r.ToString(), Font, text, rect.Left - 22,
                        rect.Top + (r - 0.5f) * size - Font.Height / 2, format);
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            var rect = GridBounds;
            if (rect.Contains(e.Location) && HoverCell != null)
                HoverCell((e.Y - rect.Y) / (rect.Height / Codebook.Rows),
                    (e.X - rect.X) / (rect.Width / Codebook.Columns));
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            var rect = GridBounds;
            if (!EditingEnabled || e.Button != MouseButtons.Left || !rect.Contains(e.Location)) return;
            int r = (e.Y - rect.Y) / (rect.Height / Codebook.Rows);
            int c = (e.X - rect.X) / (rect.Width / Codebook.Columns);
            Data[r, c] = !Data[r, c];
            Invalidate();
            if (Changed != null) Changed();
        }
    }

    public sealed class PreviewWindow : Form
    {
        public readonly MatrixView Matrix = new MatrixView();
        readonly Label summary = new Label(), source = new Label(), hover = new Label();
        readonly TextBox log = new TextBox();
        readonly Button import = new Button();
        readonly Panel right = new Panel();
        readonly TcpSession session = new TcpSession();
        readonly TextBox deviceIp = new TextBox { Text = "192.168.1.10", Dock = DockStyle.Fill };
        readonly NumericUpDown devicePort = new NumericUpDown { Minimum=1, Maximum=65535, Value=5001, Dock=DockStyle.Fill };
        readonly Button connect = new Button(), send = new Button();
        readonly CheckBox waitReply = new CheckBox { Text="等待设备 OK / ERROR 回执（5秒）", Checked=true, Dock=DockStyle.Fill };
        readonly Label connectionState = new Label { Dock=DockStyle.Fill, Text="未连接" };
        bool hasData, networking;
        bool dirty, loading;
        string sourceName = "未导入码本";

        public PreviewWindow()
        {
            Text = "全息阵面控制 · 40 行 × 64 列";
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(1200, 860);
            MinimumSize = new Size(900, 760);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(244, 247, 250);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 2, RowCount = 1 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 310));
            Controls.Add(root);
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1, Margin = new Padding(0, 0, 20, 0) };
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            root.Controls.Add(left, 0, 0);
            left.Controls.Add(new Label { Text = "数据预览", Font = new Font(Font.FontFamily, 17, FontStyle.Bold), Dock = DockStyle.Fill }, 0, 0);
            summary.Dock = DockStyle.Fill;
            left.Controls.Add(summary, 0, 1);
            left.Controls.Add(Matrix, 0, 2);
            hover.Text = "橙色 = 1（开启）    白色 = 0（关闭）";
            hover.Dock = DockStyle.Fill;
            hover.TextAlign = ContentAlignment.MiddleLeft;
            left.Controls.Add(hover, 0, 3);
            right.Dock = DockStyle.Fill;
            root.Controls.Add(right, 1, 0);
            var controls = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 15 };
            foreach (int h in new[] { 36, 42, 124, 32, 40, 40, 40, 40, 32, 28, 40, 36, 24 }) controls.RowStyles.Add(new RowStyle(SizeType.Absolute, h));
            controls.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            controls.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            right.Controls.Add(controls);
            controls.Controls.Add(new Label { Text = "设备与码本", Font = new Font(Font.FontFamily, 17, FontStyle.Bold), Dock = DockStyle.Fill }, 0, 0);
            source.Dock = DockStyle.Fill;
            source.AutoEllipsis = true;
            controls.Controls.Add(source, 0, 1);
            var device = new TableLayoutPanel { Dock=DockStyle.Fill, ColumnCount=2, RowCount=4 };
            device.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,70));
            device.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            foreach(int h in new[]{28,28,38,28}) device.RowStyles.Add(new RowStyle(SizeType.Absolute,h));
            device.Controls.Add(new Label { Text="设备 IP", Dock=DockStyle.Fill, TextAlign=ContentAlignment.MiddleLeft },0,0);
            device.Controls.Add(deviceIp,1,0);
            device.Controls.Add(new Label { Text="TCP 端口", Dock=DockStyle.Fill, TextAlign=ContentAlignment.MiddleLeft },0,1);
            device.Controls.Add(devicePort,1,1);
            SetupButton(connect,"连接设备",async delegate { await ConnectDevice(); });
            device.Controls.Add(connect,0,2); device.SetColumnSpan(connect,2);
            device.Controls.Add(connectionState,0,3); device.SetColumnSpan(connectionState,2);
            controls.Controls.Add(device,0,2);
            controls.Controls.Add(waitReply,0,3);
            waitReply.CheckedChanged += delegate {
                if(session.Connected) { session.Close(); WriteLog("回执模式已改变，请重新连接设备。"); UpdateDeviceState(); }
            };
            SetupButton(import, "导入码本…", delegate { Import(); });
            controls.Controls.Add(import, 0, 4);
            controls.Controls.Add(Button("导出 CSV…", delegate { Export(); }), 0, 5);
            controls.Controls.Add(Button("清零", delegate { if (CanReplace()) SetData(new bool[40,64], "全零码本", true); }), 0, 6);
            controls.Controls.Add(Button("加载演示图案", delegate { if (CanReplace()) SetData(Codebook.Demo(), "演示图案（非天线计算码本）", true); }), 0, 7);
            var edit = new CheckBox { Text = "允许点击网格修改 0 / 1", Dock = DockStyle.Fill };
            edit.CheckedChanged += delegate { Matrix.EditingEnabled = edit.Checked; };
            controls.Controls.Add(edit, 0, 8);
            controls.Controls.Add(new Label { Text = "逐行 · 从左到右 · 高位在前 · 320字节", Dock=DockStyle.Fill, ForeColor=Color.FromArgb(87,101,120) },0,9);
            SetupButton(send,"执行：发送当前码本",async delegate { await SendCurrent(); });
            controls.Controls.Add(send,0,10);
            controls.Controls.Add(Button("导出发送包 BIN…",delegate { ExportPacket(); }),0,11);
            controls.Controls.Add(new Label { Text = "操作记录", Dock = DockStyle.Fill }, 0, 12);
            log.Multiline = true;
            log.ReadOnly = true;
            log.ScrollBars = ScrollBars.Vertical;
            log.BackColor = Color.White;
            log.Dock = DockStyle.Fill;
            controls.Controls.Add(log, 0, 13);
            controls.Controls.Add(Button("清空记录", delegate { log.Clear(); }), 0, 14);
            Matrix.Changed += delegate { dirty = true; hasData=true; UpdateSummary(); };
            Matrix.HoverCell += delegate(int r, int c) { hover.Text = string.Format("第 {0} 行 / 第 {1} 列    值：{2}    橙色 = 1，白色 = 0", r + 1, c + 1, Matrix.Data[r,c] ? 1 : 0); };
            FormClosing += delegate(object sender, FormClosingEventArgs e) {
                if (loading) { MessageBox.Show(this, "正在读取 Excel，请等待导入完成。", "正在导入"); e.Cancel = true; }
                else if (!CanReplace()) e.Cancel = true;
                if (!e.Cancel) session.Close();
            };
            UpdateSummary();
            WriteLog("已初始化 40 行 × 64 列，共 2560 个单元。");
        }

        static void SetupButton(Button b, string text, EventHandler action)
        {
            b.Text = text; b.Dock = DockStyle.Fill; b.Margin = new Padding(0, 3, 0, 5);
            b.FlatStyle = FlatStyle.Flat; b.BackColor = Color.White;
            b.FlatAppearance.BorderColor = Color.FromArgb(196, 205, 216); b.Click += action;
        }
        static Button Button(string text, EventHandler action) { var b = new Button(); SetupButton(b, text, action); return b; }
        bool CanReplace() { return !dirty || MessageBox.Show(this, "当前修改尚未导出，继续将放弃这些修改。", "未导出的修改", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK; }
        void UpdateSummary()
        {
            int count = Matrix.Data.Cast<bool>().Count(x => x);
            summary.Text = "40 行 × 64 列  /  2560 单元    开启 " + count + "    关闭 " + (2560 - count);
            source.Text = "当前数据：" + sourceName + (dirty ? "\n有未导出的修改" : "");
            UpdateDeviceState();
        }
        public void SetData(bool[,] data, string name, bool modified)
        {
            if (data.GetLength(0) != 40 || data.GetLength(1) != 64) throw new ArgumentException("需要 40 × 64 数据");
            Matrix.Data = (bool[,])data.Clone(); sourceName = name; dirty = modified; hasData=true;
            UpdateSummary(); Matrix.Invalidate(); WriteLog("已显示：" + name);
        }
        void WriteLog(string text) { log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + text + Environment.NewLine); }
        void UpdateDeviceState()
        {
            send.Enabled = hasData && session.Connected && !networking && !loading;
            deviceIp.Enabled = devicePort.Enabled = !session.Connected;
            connect.Text = session.Connected ? "断开连接" : "连接设备";
            connectionState.Text = session.Connected ? "已连接 " + deviceIp.Text + ":" + devicePort.Value : "未连接";
        }
        async System.Threading.Tasks.Task ConnectDevice()
        {
            if(session.Connected) { session.Close(); WriteLog("连接已断开。"); UpdateDeviceState(); return; }
            networking=true; right.Enabled=false;
            string ip=deviceIp.Text.Trim(); int port=(int)devicePort.Value;
            WriteLog("正在连接 TCP " + ip + ":" + port);
            try { await session.ConnectAsync(ip,port,3000); if(!IsDisposed) WriteLog("连接成功。"); }
            catch(Exception ex) { if(!IsDisposed) WriteLog("连接失败："+ex.Message); }
            finally { networking=false; if(!IsDisposed) { right.Enabled=true; UpdateDeviceState(); } }
        }
        async System.Threading.Tasks.Task SendCurrent()
        {
            if(!hasData || networking || loading) return;
            byte[] packet=PacketEncoder.Encode((bool[,])Matrix.Data.Clone(),BitOrder.MostSignificantFirst);
            bool requireReply=waitReply.Checked;
            networking=true; right.Enabled=false; Matrix.Enabled=false;
            WriteLog("发送当前码本：320 字节；"+(requireReply?"等待设备回执。":"不等待回执。"));
            try
            {
                string reply=await session.SendAsync(packet,requireReply,5000);
                if(!IsDisposed) WriteLog(reply=="OK" ? "设备返回 OK：执行成功。" :
                    reply=="ERROR" ? "设备返回 ERROR：设备拒绝数据。" : "320 字节已提交发送；未等待回执，执行结果未确认。");
            }
            catch(Exception ex) { if(!IsDisposed) WriteLog("发送未确认："+ex.Message); }
            finally { networking=false; if(!IsDisposed) { right.Enabled=true; Matrix.Enabled=true; UpdateDeviceState(); } }
        }
        void ExportPacket()
        {
            if(!hasData) { MessageBox.Show(this,"请先导入、编辑或清零生成码本。","尚无码本"); return; }
            using(var dialog=new SaveFileDialog { Filter="二进制发送包|*.bin", FileName="codebook_40x64.bin" })
            {
                if(dialog.ShowDialog(this)!=DialogResult.OK) return;
                try { File.WriteAllBytes(dialog.FileName,PacketEncoder.Encode(Matrix.Data,BitOrder.MostSignificantFirst)); WriteLog("已导出 320 字节发送包："+dialog.FileName); }
                catch(Exception ex) { MessageBox.Show(this,ex.Message,"导出失败"); }
            }
        }
        void Import()
        {
            if (!CanReplace()) return;
            using (var dialog = new OpenFileDialog { Filter = "码本文件|*.xls;*.xlsx;*.csv;*.tsv;*.txt", Title = "导入 40 行 × 64 列码本" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                string path = dialog.FileName;
                loading = true; right.Enabled = false; Matrix.Enabled = false;
                WriteLog("正在读取：" + Path.GetFileName(path));
                var worker = new Thread(delegate() {
                    bool[,] data = null; Exception error = null;
                    try { data = Codebook.Read(path); } catch (Exception ex) { error = ex; }
                    BeginInvoke((Action)delegate {
                        loading = false; right.Enabled = true; Matrix.Enabled = true; UpdateDeviceState();
                        if (error != null) { WriteLog("导入失败：" + error.Message); MessageBox.Show(this, error.Message, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                        else SetData(data, Path.GetFileName(path), false);
                    });
                });
                worker.IsBackground = true; worker.SetApartmentState(ApartmentState.STA); worker.Start();
            }
        }
        void Export()
        {
            using (var dialog = new SaveFileDialog { Filter = "CSV 码本|*.csv", FileName = "codebook_40x64.csv" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try { Codebook.Save(dialog.FileName, Matrix.Data); dirty = false; sourceName = Path.GetFileName(dialog.FileName); UpdateSummary(); WriteLog("已导出：" + dialog.FileName); }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, "导出失败"); }
            }
        }
    }

    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length > 0 && args[0] == "--self-test") return SelfTest(args[1]);
            Application.Run(new PreviewWindow());
            return 0;
        }
        static void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
        static int SelfTest(string output)
        {
            Directory.CreateDirectory(output);
            var report = new List<string>();
            try
            {
                var original = Codebook.Demo();
                string csv = Path.Combine(output, "demo_40x64.csv");
                Codebook.Save(csv, original);
                var loaded = Codebook.Read(csv);
                Assert(original.Cast<bool>().SequenceEqual(loaded.Cast<bool>()), "CSV round trip");
                report.Add("PASS: 40x64 CSV import/export round trip, 2560 cells");
                var invalid = Enumerable.Range(0,40).Select(r => (IList<string>)Enumerable.Repeat("0",64).ToArray()).ToList();
                invalid[39][63] = "2";
                try { Codebook.Validate(invalid); throw new Exception("invalid accepted"); } catch (FormatException) { }
                invalid[39][63] = "";
                try { Codebook.Validate(invalid); throw new Exception("blank accepted"); } catch (FormatException) { }
                invalid.RemoveAt(39);
                try { Codebook.Validate(invalid); throw new Exception("wrong dimensions accepted"); } catch (FormatException) { }
                report.Add("PASS: non-binary, blank and wrong dimensions rejected");
                using (var form = new PreviewWindow())
                {
                    form.SetData(original, "演示图案（非天线计算码本）", false);
                    foreach (var size in new[] { new Size(1200,860), new Size(900,760), new Size(1500,960) })
                    {
                        form.Size = size; form.CreateControl(); form.PerformLayout();
                        // Render the root without an invisible Form ancestor, so WinForms
                        // includes child controls while no desktop window is opened.
                        Control root = form.Controls[0];
                        Size rootSize = root.Size;
                        form.Controls.Remove(root);
                        root.Dock = DockStyle.None;
                        root.Size = rootSize;
                        root.CreateControl(); root.PerformLayout();
                        using (var bitmap = new Bitmap(root.Width, root.Height))
                        {
                            root.DrawToBitmap(bitmap, root.ClientRectangle);
                            bitmap.Save(Path.Combine(output, "preview_" + size.Width + ".png"));
                        }
                        var grid = form.Matrix.GridBounds;
                        Assert(grid.Width / 64 == grid.Height / 40, "cells not square");
                        Assert(form.Matrix.ClientRectangle.Contains(grid), "grid clipped");
                        form.Controls.Add(root); root.Dock = DockStyle.Fill;
                    }
                    using (var bitmap = new Bitmap(form.Matrix.Width, form.Matrix.Height))
                    {
                        form.Matrix.DrawToBitmap(bitmap, form.Matrix.ClientRectangle);
                        var rect = form.Matrix.GridBounds;
                        int cell = rect.Width / 64;
                        for (int r=0;r<40;r++) for(int c=0;c<64;c++)
                        {
                            Color actual = bitmap.GetPixel(rect.X+c*cell+cell/2,rect.Y+r*cell+cell/2);
                            Color expected = original[r,c] ? Color.FromArgb(255,111,35) : Color.White;
                            Assert(actual.ToArgb()==expected.ToArgb(), "pixel mismatch at "+r+","+c);
                        }
                    }
                    report.Add("PASS: all 2560 cell colors and positions; three window sizes without clipping");
                    var click = typeof(MatrixView).GetMethod("OnMouseClick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    var bounds = form.Matrix.GridBounds;
                    int cellWidth = bounds.Width / 64;
                    var bottomRight = new MouseEventArgs(MouseButtons.Left, 1, bounds.Right-cellWidth/2-1, bounds.Bottom-cellWidth/2-1, 0);
                    bool prior = form.Matrix.Data[39,63];
                    click.Invoke(form.Matrix, new object[] { bottomRight });
                    Assert(form.Matrix.Data[39,63] == prior, "editing should be disabled initially");
                    form.Matrix.EditingEnabled = true;
                    click.Invoke(form.Matrix, new object[] { bottomRight });
                    Assert(form.Matrix.Data[39,63] != prior, "last cell click mapping");
                    click.Invoke(form.Matrix, new object[] { new MouseEventArgs(MouseButtons.Left,1,bounds.Right,bounds.Bottom,0) });
                    Assert(form.Matrix.Data[39,63] != prior, "outside click should do nothing");
                    report.Add("PASS: edit toggle, bottom-right cell hit test, outside-grid click");
                    Assert(logHeight(form)>=60,"operation log clipped at minimum window size");
                    TestDeviceButtons(form);
                    report.Add("PASS: UI defaults 192.168.1.10:5001, connect/send/disconnect buttons, exact preview payload and OK log");
                }
                File.WriteAllLines(Path.Combine(output,"test-results.txt"),report);
                return 0;
            }
            catch(Exception ex) { report.Add("FAIL: "+ex); File.WriteAllLines(Path.Combine(output,"test-results.txt"),report); return 1; }
        }
        static object Field(PreviewWindow form,string name)
        {
            return typeof(PreviewWindow).GetField(name,System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic).GetValue(form);
        }
        static int logHeight(PreviewWindow form) { return ((TextBox)Field(form,"log")).ClientSize.Height; }
        static void Pump(System.Threading.Tasks.Task task)
        {
            var time=System.Diagnostics.Stopwatch.StartNew();
            while(!task.IsCompleted && time.ElapsedMilliseconds<10000) { Application.DoEvents(); Thread.Sleep(5); }
            Assert(task.IsCompleted,"UI task did not complete"); task.GetAwaiter().GetResult();
        }
        static void TestDeviceButtons(PreviewWindow form)
        {
            var ip=(TextBox)Field(form,"deviceIp");
            var port=(NumericUpDown)Field(form,"devicePort");
            var send=(Button)Field(form,"send");
            var session=(TcpSession)Field(form,"session");
            Assert(ip.Text=="192.168.1.10" && port.Value==5001,"device defaults incorrect");
            Assert(!send.Enabled,"disconnected send button enabled");
            var listener=new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback,0);
            listener.Start();
            byte[] expected=PacketEncoder.Encode(form.Matrix.Data,BitOrder.MostSignificantFirst);
            var peer=System.Threading.Tasks.Task.Run(async delegate {
                using(var accepted=await listener.AcceptTcpClientAsync())
                {
                    var stream=accepted.GetStream(); byte[] actual=new byte[320]; int offset=0;
                    while(offset<320) { int count=await stream.ReadAsync(actual,offset,320-offset); if(count==0) throw new Exception("UI short payload"); offset+=count; }
                    Assert(actual.SequenceEqual(expected),"UI sends different data from grid");
                    byte[] reply=Encoding.ASCII.GetBytes("OK"); await stream.WriteAsync(reply,0,reply.Length);
                    await stream.ReadAsync(new byte[1],0,1);
                }
            });
            try
            {
                ip.Text="127.0.0.1"; port.Value=((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
                var flags=System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic;
                var connect=typeof(PreviewWindow).GetMethod("ConnectDevice",flags);
                Pump((System.Threading.Tasks.Task)connect.Invoke(form,null));
                Assert(session.Connected && send.Enabled,"UI connection did not enable send");
                Pump((System.Threading.Tasks.Task)typeof(PreviewWindow).GetMethod("SendCurrent",flags).Invoke(form,null));
                Assert(((TextBox)Field(form,"log")).Text.Contains("设备返回 OK：执行成功。"),"UI missing execution acknowledgement");
                Pump((System.Threading.Tasks.Task)connect.Invoke(form,null));
                Assert(!session.Connected && !send.Enabled,"UI disconnect leaves send enabled");
                Pump(peer);
            }
            finally { session.Close(); listener.Stop(); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultiToolWin.Pages
{
    /// <summary>
    /// 简版：按 Excel A 列（或 CSV/TXT 每行）批量创建文件夹
    /// - 不做名称清洗；非法/保留名将创建失败并记录日志
    /// - 已存在则跳过
    /// - 支持开始/停止、进度与日志
    /// </summary>
    public class PageMkFolders : UserControl
    {
        private readonly Action<string> Log;

        // 输入
        private TextBox txtExcel, txtParent;

        // 控件
        private Button btnStart, btnStop;

        // 选项
        private CheckBox chkOpenWhenDone;

        // 进度显示 / 取消
        private ProgressBar pb;
        private Label lblCurrent, lblCounts, lblElapsed;
        private System.Diagnostics.Stopwatch _sw;
        private System.Windows.Forms.Timer _uiTimer;
        private CancellationTokenSource _cts;

        public PageMkFolders(Action<string> logger)
        {
            Log = logger;
            BuildUI();
        }

        private void BuildUI()
        {
            this.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
            const int LABEL_COL_WIDTH = 100;
            const int BUTTON_COL_WIDTH = 88;

            var rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(12),
                AutoScroll = true
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var filesGroup = new GroupBox
            {
                Text = "文件与目录",
                Dock = DockStyle.Top,
                Height = 132,
                Padding = new Padding(8),
                Margin = new Padding(0, 0, 0, 8)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 6,
                Padding = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, LABEL_COL_WIDTH));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, BUTTON_COL_WIDTH));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, BUTTON_COL_WIDTH));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 0: Excel
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1: 新建文件目录
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2: 选项
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3: 进度区
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 4: 按钮行
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 5: 占位

            // 小工厂
            Label MkLabel(string text) => new Label
            {
                Text = text,
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                Margin = new Padding(0, 6, 8, 6)
            };
            TextBox MkText() => new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 6, 3)
            };
            Button MkBtn(string text)
            {
                return new Button
                {
                    Text = text,
                    AutoSize = false,
                    Width = BUTTON_COL_WIDTH,
                    Height = 26,                    // 统一高度
                    Anchor = AnchorStyles.Right,     // 在单元格内靠右
                    Margin = new Padding(0, 3, 0, 3),
                    FlatStyle = FlatStyle.System
                };
            }


            // ① Excel 文件
            layout.Controls.Add(MkLabel("Excel 文件："), 0, 0);
            txtExcel = MkText();
            layout.Controls.Add(txtExcel, 1, 0);
            var btnPickExcel = MkBtn("浏览…");
            btnPickExcel.Click += (s, e) =>
            {
                var ofd = new OpenFileDialog
                {
                    Filter = "Excel/CSV/TXT|*.xlsx;*.xls;*.csv;*.txt|所有文件|*.*"
                };
                if (ofd.ShowDialog() == DialogResult.OK) txtExcel.Text = ofd.FileName;
            };
            layout.Controls.Add(btnPickExcel, 2, 0);
            layout.Controls.Add(new Panel { Width = 1, Height = 1 }, 3, 0); // 占位

            // ② 新建文件目录
            layout.Controls.Add(MkLabel("新建文件目录："), 0, 1);
            txtParent = MkText();
            layout.Controls.Add(txtParent, 1, 1);
            var btnPickParent = MkBtn("浏览…");
            btnPickParent.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog())
                    if (fbd.ShowDialog() == DialogResult.OK) txtParent.Text = fbd.SelectedPath;
            };
            layout.Controls.Add(btnPickParent, 2, 1);
            var btnOpenParent = MkBtn("打开…");
            btnOpenParent.Click += (s, e) =>
            {
                var p = (txtParent.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(p) || !Directory.Exists(p))
                {
                    MessageBox.Show("请先选择有效的新建文件目录。", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                try { System.Diagnostics.Process.Start("explorer.exe", p); }
                catch { Log?.Invoke("打开新建文件目录失败。"); }
            };
            layout.Controls.Add(btnOpenParent, 3, 1);

            // ③ 选项
            var panelOptions = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 6, 0, 0)
            };
            chkOpenWhenDone = new CheckBox { Text = "完成后打开新建文件目录", Checked = true, AutoSize = true, Anchor = AnchorStyles.Left };
            panelOptions.Controls.Add(chkOpenWhenDone);
            layout.Controls.Add(panelOptions, 0, 2);
            layout.SetColumnSpan(panelOptions, 4);

            // ④ 进度区
            var progressGroup = new GroupBox { Text = "执行进度", Dock = DockStyle.Top, Height = 118, Padding = new Padding(8), Margin = new Padding(0, 6, 0, 0) };
            var panelProgress = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            panelProgress.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            panelProgress.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panelProgress.Controls.Add(new Label { Text = "当前：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            lblCurrent = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblCurrent, 1, 0);
            panelProgress.Controls.Add(new Label { Text = "进度：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            pb = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100, Value = 0 };
            panelProgress.Controls.Add(pb, 1, 1);
            panelProgress.Controls.Add(new Label { Text = "计数：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            lblCounts = new Label { Text = "已创建 0｜已存在 0｜失败 0", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblCounts, 1, 2);
            panelProgress.Controls.Add(new Label { Text = "用时：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            lblElapsed = new Label { Text = "00:00:00", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblElapsed, 1, 3);
            progressGroup.Controls.Add(panelProgress);

            // ⑤ 按钮行
            btnStart = new Button { Text = "开始创建", AutoSize = false, Width = BUTTON_COL_WIDTH, Height = 28, Margin = new Padding(6, 3, 0, 3), FlatStyle = FlatStyle.System };
            btnStop = new Button { Text = "停止", AutoSize = false, Width = BUTTON_COL_WIDTH, Height = 28, Margin = new Padding(0, 3, 0, 3), Enabled = false, FlatStyle = FlatStyle.System };

            btnStart.Click += (s, e) => DoMakeFolders();
            btnStop.Click += (s, e) => _cts?.Cancel();

            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 6, 0, 0)
            };
            btnPanel.Controls.Add(btnStop);
            btnPanel.Controls.Add(btnStart);
            filesGroup.Controls.Add(layout);
            rootLayout.Controls.Add(filesGroup, 0, 0);
            rootLayout.Controls.Add(progressGroup, 0, 1);
            rootLayout.Controls.Add(btnPanel, 0, 2);

            Controls.Add(rootLayout);
        }

        // 开始/结束 & 进度
        private void BeginWork()
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            _cts = new CancellationTokenSource();

            _sw = System.Diagnostics.Stopwatch.StartNew();
            _uiTimer = new System.Windows.Forms.Timer { Interval = 500 };
            _uiTimer.Tick += (s, e) =>
            {
                if (_sw != null) lblElapsed.Text = _sw.Elapsed.ToString(@"hh\:mm\:ss");
            };
            _uiTimer.Start();
        }

        private void EndWork()
        {
            try { _cts?.Cancel(); } catch { }
            _cts = null;

            try { _uiTimer?.Stop(); _uiTimer?.Dispose(); } catch { }
            _uiTimer = null;

            try { _sw?.Stop(); } catch { }
            _sw = null;

            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        private void UpdateProgressSafe(int done, int total, string current, int created, int existed, int failed)
        {
            if (IsHandleCreated)
            {
                BeginInvoke((Action)(() =>
                {
                    lblCurrent.Text = string.IsNullOrEmpty(current) ? "-" : current;
                    pb.Value = total > 0 ? Math.Min(100, (int)Math.Round(done * 100.0 / total)) : 0;
                    lblCounts.Text = $"已创建 {created}｜已存在 {existed}｜失败 {failed}";
                }));
            }
        }

        // 主流程：按 A 列创建文件夹（不清洗）
        private void DoMakeFolders()
        {
            var excel = (txtExcel.Text ?? "").Trim();
            var parent = (txtParent.Text ?? "").Trim();

            if (!File.Exists(excel)) { Log?.Invoke("请选择有效的 Excel/CSV/TXT 文件。"); return; }
            if (string.IsNullOrWhiteSpace(parent) || !Directory.Exists(parent)) { Log?.Invoke("请选择有效的新建文件目录。"); return; }

            // 读取名称列表（A 列 / 每行）
            List<string> names;
            try
            {
                names = ReadNamesFromExcelOrText(excel);
            }
            catch (Exception ex)
            {
                Log?.Invoke($"读取失败：{ex.Message}");
                return;
            }

            // 过滤空白 & 去重（忽略大小写）
            var valid = names
                        .Select(s => (s ?? "").Trim())
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

            if (valid.Count == 0) { Log?.Invoke("Excel/CSV/TXT 中未找到有效的名称。"); return; }

            Log?.Invoke($"将创建目录数（去重后）：{valid.Count}");
            BeginWork();

            Task.Run(() =>
            {
                int total = valid.Count;
                int done = 0, created = 0, existed = 0, failed = 0;

                foreach (var name in valid)
                {
                    if (_cts?.IsCancellationRequested == true) { Log?.Invoke("已停止。"); break; }

                    var dest = Path.Combine(parent, name);
                    string current = name;

                    try
                    {
                        if (Directory.Exists(dest))
                        {
                            existed++;
                            Log?.Invoke($"已存在，跳过：{dest}");
                        }
                        else
                        {
                            Directory.CreateDirectory(dest);
                            created++;
                            Log?.Invoke($"创建：{dest}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        Log?.Invoke($"失败：{dest} —— {ex.Message}");
                    }

                    done++;
                    UpdateProgressSafe(done, total, current, created, existed, failed);
                }

                Log?.Invoke($"完成：共 {total}，已创建 {created}，已存在 {existed}，失败 {failed}");

                if (chkOpenWhenDone.Checked)
                {
                    try { System.Diagnostics.Process.Start("explorer.exe", parent); } catch { }
                }
            })
            .ContinueWith(_ => { try { EndWork(); } catch { } }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        /// <summary>
        /// 读取名称列表：
        /// - *.csv/*.txt：逐行读取
        /// - 其他（如 xlsx/xls）：尝试通过 ExcelUtils 反射调用：
        ///   - MultiToolWin.Utils.ExcelUtils.ReadFolderNames(string)
        ///   - 或 ExcelUtils.ReadFirstColumn(string)
        ///   如未找到方法则抛出异常，提示实现
        /// </summary>
        private List<string> ReadNamesFromExcelOrText(string path)
        {
            var ext = Path.GetExtension(path)?.ToLowerInvariant();
            if (ext == ".csv" || ext == ".txt")
            {
                // 文本/CSV：每行一个名称
                return File.ReadAllLines(path)
                           .Select(l => l?.Trim() ?? "")
                           .ToList();
            }

            // 反射尝试调用 ExcelUtils
            try
            {
                var utilsType = AppDomain.CurrentDomain
                    .GetAssemblies()
                    .SelectMany(a =>
                    {
                        try { return a.GetTypes(); } catch { return Array.Empty<Type>(); }
                    })
                    .FirstOrDefault(t => t.FullName == "MultiToolWin.Utils.ExcelUtils");

                if (utilsType != null)
                {
                    // 优先 ReadFolderNames(string)
                    var m1 = utilsType.GetMethod("ReadFolderNames", new[] { typeof(string) });
                    if (m1 != null)
                    {
                        var result = m1.Invoke(null, new object[] { path }) as System.Collections.IEnumerable;
                        return result?.Cast<object>().Select(o => o?.ToString() ?? "").ToList() ?? new List<string>();
                    }
                    // 备选 ReadFirstColumn(string)
                    var m2 = utilsType.GetMethod("ReadFirstColumn", new[] { typeof(string) });
                    if (m2 != null)
                    {
                        var result = m2.Invoke(null, new object[] { path }) as System.Collections.IEnumerable;
                        return result?.Cast<object>().Select(o => o?.ToString() ?? "").ToList() ?? new List<string>();
                    }
                }
            }
            catch { /* 忽略反射异常，走提示 */ }

            throw new NotSupportedException(
                "未找到 Excel 读取方法。请在 ExcelUtils 中实现 ReadFolderNames(string) 或 ReadFirstColumn(string)。");
        }
    }
}

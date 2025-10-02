using MultiToolWin.Utils;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultiToolWin.Pages
{
    public class PageExtract : UserControl
    {
        private Action<string> Log;

        // —— 输入与选项 —— 
        private TextBox txtExcel, txtInput, txtOutput;
        private RadioButton rbCopy, rbMove;

        // —— 按钮 —— 
        private Button btnStart, btnStop;

        // —— 进度显示 / 取消 —— 
        private ProgressBar pb;
        private Label lblCurrent, lblCounts, lblElapsed;
        private System.Diagnostics.Stopwatch _sw;
        private System.Windows.Forms.Timer _uiTimer;
        private CancellationTokenSource _cts;

        // —— 文件日志（批量落盘，避免卡顿） —— 
        private CheckBox chkSaveLog;
        private Label lblLogPath;
        private string _logFilePath = null;

        private readonly ConcurrentQueue<string> _logQueue = new ConcurrentQueue<string>();
        private volatile bool _logWriterRunning = false;
        private Task _logWriterTask;
        private StreamWriter _logWriter;
        private DateTime _lastUiLog = DateTime.MinValue; // UI 日志限频

        public PageExtract(Action<string> logger)
        {
            Log = logger;
            BuildUI();
        }

        private void BuildUI()
        {
            // —— 列宽常量（可微调）——
            int LABEL_COL_WIDTH = 100;  // 第0列标签固定宽
            int BUTTON_COL_WIDTH = 88;  // “浏览/打开/开始/停止”统一宽

            // 4列：标签 | 文本框 | 浏览按钮 | 打开按钮
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 7,
                Padding = new Padding(12)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, LABEL_COL_WIDTH));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, BUTTON_COL_WIDTH));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, BUTTON_COL_WIDTH));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 0: Excel
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1: 输入
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2: 输出
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3: 操作方式
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 4: 选项区
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 5: 进度区
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 6: 按钮行

            // —— 小工厂 —— 
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
                var b = new Button
                {
                    Text = text,
                    AutoSize = false,
                    Width = BUTTON_COL_WIDTH,   // 保持你原来的常量
                    Height = 25,                // ← 新增：统一高度
                    Anchor = AnchorStyles.Right, // ← 新增：在单元格里靠右
                    Margin = new Padding(0, 3, 6, 3), // ← 上下留点间距，看起来居中
                    FlatStyle = FlatStyle.System
                };
                return b;
            } // ← 小工厂结束
              // ← 小工厂结束；从这里开始正式组装 UI

            // ① Excel
            layout.Controls.Add(MkLabel("Excel 文件："), 0, 0);
            txtExcel = MkText();
            layout.Controls.Add(txtExcel, 1, 0);
            var btnExcel = MkBtn("浏览…");
            btnExcel.Click += (s, e) =>
            {
                var ofd = new OpenFileDialog { Filter = "Excel 文件|*.xlsx;*.xls" };
                if (ofd.ShowDialog() == DialogResult.OK) txtExcel.Text = ofd.FileName;
            };
            layout.Controls.Add(btnExcel, 2, 0);
            layout.Controls.Add(new Panel { Width = 1, Height = 1 }, 3, 0); // 占位

            // ② 输入
            layout.Controls.Add(MkLabel("输入文件夹："), 0, 1);
            txtInput = MkText();
            layout.Controls.Add(txtInput, 1, 1);
            var btnPickInput = MkBtn("浏览…");
            btnPickInput.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog())
                    if (fbd.ShowDialog() == DialogResult.OK) txtInput.Text = fbd.SelectedPath;
            };
            layout.Controls.Add(btnPickInput, 2, 1);
            layout.Controls.Add(new Panel { Width = 1, Height = 1 }, 3, 1); // 占位

            // ③ 输出
            layout.Controls.Add(MkLabel("输出目录："), 0, 2);
            txtOutput = MkText();
            layout.Controls.Add(txtOutput, 1, 2);
            var btnPickOut = MkBtn("浏览…");
            btnPickOut.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog())
                    if (fbd.ShowDialog() == DialogResult.OK) txtOutput.Text = fbd.SelectedPath;
            };
            layout.Controls.Add(btnPickOut, 2, 2);
            var btnOpenOut = MkBtn("打开…");
            btnOpenOut.Click += (s, e) =>
            {
                // 留空则打开“输入根”，否则打开“输出目录”
                var p = (txtOutput.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(p)) p = (txtInput.Text ?? "").Trim();
                if (string.IsNullOrWhiteSpace(p) || !Directory.Exists(p))
                {
                    MessageBox.Show("请先选择有效的目录。", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                try { System.Diagnostics.Process.Start("explorer.exe", p); }
                catch { Log?.Invoke("打开目录失败。"); }
            };
            layout.Controls.Add(btnOpenOut, 3, 2);

            // ④ 操作方式
            layout.Controls.Add(MkLabel("操作方式："), 0, 3);
            var ops = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 3, 0, 3),
                WrapContents = false
            };
            rbCopy = new RadioButton { Text = "复制", AutoSize = true, Margin = new Padding(0, 0, 18, 0) };
            rbMove = new RadioButton { Text = "剪切", AutoSize = true };
            rbMove.Checked = true; // 如需默认“复制”，把这行去掉
            ops.Controls.Add(rbCopy);
            ops.Controls.Add(rbMove);
            layout.Controls.Add(ops, 1, 3);
            layout.SetColumnSpan(ops, 3);

            // ⑤ 选项区：保存日志到文件
            var panelOptions = new TableLayoutPanel { Dock = DockStyle.Top, Height = 60, ColumnCount = 2, Padding = new Padding(0, 6, 0, 0) };
            panelOptions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            panelOptions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            chkSaveLog = new CheckBox { Text = "保存日志到文件", AutoSize = true, Anchor = AnchorStyles.Left };
            chkSaveLog.CheckedChanged += (s, e) =>
            {
                if (chkSaveLog.Checked)
                {
                    using (var sfd = new SaveFileDialog
                    {
                        Filter = "文本文件|*.txt|所有文件|*.*",
                        FileName = "提取日志.txt",
                        InitialDirectory = Directory.Exists(txtOutput.Text) ? txtOutput.Text
                                          : (Directory.Exists(txtInput.Text) ? txtInput.Text
                                          : AppDomain.CurrentDomain.BaseDirectory)
                    })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            _logFilePath = sfd.FileName;
                            lblLogPath.Text = $"日志路径：{_logFilePath}";
                            lblLogPath.Visible = true;
                            try
                            {
                                if (!File.Exists(_logFilePath))
                                    File.WriteAllText(_logFilePath, $"[创建] {DateTime.Now:yyyy-MM-dd HH:mm:ss} 日志开始\r\n");
                            }
                            catch (Exception ex)
                            {
                                _logFilePath = null;
                                lblLogPath.Visible = false;
                                chkSaveLog.Checked = false;
                                Log?.Invoke($"创建日志文件失败：{ex.Message}");
                            }
                        }
                        else
                        {
                            chkSaveLog.Checked = false;
                        }
                    }
                }
                else
                {
                    _logFilePath = null;
                    lblLogPath.Visible = false;
                }
            };
            lblLogPath = new Label { AutoSize = true, Visible = false, Anchor = AnchorStyles.Left, Padding = new Padding(12, 6, 0, 0) };
            layout.Controls.Add(panelOptions, 0, 4);
            layout.SetColumnSpan(panelOptions, 4);
            panelOptions.Controls.Add(chkSaveLog, 0, 0);
            panelOptions.Controls.Add(lblLogPath, 1, 0);

            // ⑥ 进度区
            var panelProgress = new TableLayoutPanel { Dock = DockStyle.Top, Height = 90, ColumnCount = 2, Padding = new Padding(0, 6, 0, 0) };
            panelProgress.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            panelProgress.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panelProgress.Controls.Add(new Label { Text = "当前：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            lblCurrent = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblCurrent, 1, 0);
            panelProgress.Controls.Add(new Label { Text = "进度：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            pb = new ProgressBar { Dock = DockStyle.Fill, Minimum = 0, Maximum = 100, Value = 0 };
            panelProgress.Controls.Add(pb, 1, 1);
            panelProgress.Controls.Add(new Label { Text = "计数：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            lblCounts = new Label { Text = "0 / 0 | 成功 0 失败 0 缺数 0 未找到 0", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblCounts, 1, 2);
            panelProgress.Controls.Add(new Label { Text = "用时：", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            lblElapsed = new Label { Text = "00:00:00", AutoSize = true, Anchor = AnchorStyles.Left };
            panelProgress.Controls.Add(lblElapsed, 1, 3);
            layout.Controls.Add(panelProgress, 0, 5);
            layout.SetColumnSpan(panelProgress, 4);

            // ⑦ 按钮行
            btnStart = new Button
            {
                Text = "提取",
                AutoSize = false,
                Width = 80,
                Height = 28,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 3, 10, 3), // ← 右边多留10px 空隙
                FlatStyle = FlatStyle.System

            };
            btnStop = new Button
            {
                Text = "停止",
                AutoSize = false,
                Width = 80,
                Height = 28,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 3, 10, 3), // ← 右边多留10px 空隙
                FlatStyle = FlatStyle.System
            };
            btnStart.Click += (s, e) => DoExtract();
            btnStop.Click += (s, e) => _cts?.Cancel();

            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 6, 0, 0)
            };
            btnPanel.Controls.Add(btnStop);
            btnPanel.Controls.Add(btnStart);
            layout.Controls.Add(btnPanel, 0, 6);
            layout.SetColumnSpan(btnPanel, 4);

            // 装载
            Controls.Add(layout);
        }

        // —— 启动/停止日志落盘（选了路径才会启用） —— 
        private void StartLogWriterIfNeeded()
        {
            if (_logWriterRunning) return;

            if (!string.IsNullOrEmpty(_logFilePath))
            {
                try { _logWriter = new StreamWriter(_logFilePath, append: true, System.Text.Encoding.UTF8); }
                catch (Exception ex)
                {
                    _logWriter = null;
                    Log?.Invoke($"打开日志文件失败：{ex.Message}");
                }
            }

            _logWriterRunning = true;
            _logWriterTask = Task.Run(async () =>
            {
                var batch = new List<string>(256);
                while (_logWriterRunning)
                {
                    try
                    {
                        batch.Clear();
                        while (_logQueue.TryDequeue(out var line))
                        {
                            batch.Add(line);
                            if (batch.Count >= 256) break;   // 批量上限
                        }
                        if (batch.Count > 0 && _logWriter != null)
                        {
                            foreach (var l in batch) _logWriter.WriteLine(l);
                            _logWriter.Flush();
                        }
                    }
                    catch { /* 忽略单次落盘异常 */ }
                    await Task.Delay(400);
                }
            });
        }

        private void StopLogWriter()
        {
            _logWriterRunning = false;
            try { _logWriterTask?.Wait(1000); } catch { }
            try { _logWriter?.Flush(); _logWriter?.Dispose(); } catch { }
            _logWriterTask = null;
            _logWriter = null;
        }

        // —— 开始/结束一个长任务（统一切换 UI 与计时/日志后台） —— 
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

            StartLogWriterIfNeeded();
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

            StopLogWriter();
        }

        // —— UI 线程安全的进度更新 —— 
        private void UpdateProgressSafe(int done, int total, string current, int ok, int fail, int lack, int miss)
        {
            if (IsHandleCreated)
            {
                BeginInvoke((Action)(() =>
                {
                    lblCurrent.Text = string.IsNullOrEmpty(current) ? "-" : current;
                    pb.Value = total > 0 ? Math.Min(100, (int)Math.Round(done * 100.0 / total)) : 0;
                    lblCounts.Text = $"{done} / {total} | 成功 {ok} 失败 {fail} 缺数 {lack} 未找到 {miss}";
                }));
            }
        }

        // —— UI 限频 + 文件日志入队（窗口“仅保留10条”由外部 Log 实现） —— 
        private void LogEx(string message)
        {
            var now = DateTime.UtcNow;
            if ((now - _lastUiLog).TotalMilliseconds >= 80
                || message.StartsWith("完成") || message.StartsWith("失败") || message.StartsWith("成功"))
            {
                _lastUiLog = now;
                try { Log?.Invoke(message); } catch { }
            }

            if (!string.IsNullOrEmpty(_logFilePath))
            {
                _logQueue.Enqueue($"[{DateTime.Now:HH:mm:ss}] {message}");
            }
        }

        // ========= 按 Excel A/B/C 三列提取 =========
        // 规则改动：当 C<=0 或为空 → 只创建目标文件夹，不提取图片
        // 输出目录留空 → 直接放到各自源文件夹（A列）下面的新文件夹（B列）中
        private void DoExtract()
        {
            var excel = txtExcel.Text?.Trim();
            var input = txtInput.Text?.Trim();
            var outputRoot = (txtOutput.Text ?? "").Trim();
            bool useSourceAsOutput = string.IsNullOrWhiteSpace(outputRoot);

            if (!File.Exists(excel)) { LogEx("请选择有效的 Excel 文件。"); return; }
            if (string.IsNullOrWhiteSpace(input) || !Directory.Exists(input)) { LogEx("请选择有效的输入文件夹。"); return; }

            if (!useSourceAsOutput)
            {
                try { Directory.CreateDirectory(outputRoot); } catch { /* 忽略 */ }
                LogEx($"输出目录（固定）：{outputRoot}");
            }
            else
            {
                LogEx("输出目录：留空 → 每个源文件夹内部创建目标子文件夹。");
            }

            var plan = ExcelUtils.ReadExtractPlan(excel); // 识别：原文件夹名字 / 新文件夹名字 / 页数
            if (plan.Count == 0) { LogEx("Excel 中未找到“原文件夹名字/新文件夹名字”列，或无数据。"); return; }

            var allow = new HashSet<string>(new[] { ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".pdf" },
                                            StringComparer.OrdinalIgnoreCase);

            // —— 预估总量（用于进度条）：C<=0 时不计入 —— 
            int total = 0;
            try
            {
                foreach (var grp in plan.GroupBy(p => p.Src, StringComparer.OrdinalIgnoreCase))
                {
                    var srcDir = Path.Combine(input, grp.Key);
                    if (!Directory.Exists(srcDir)) continue;

                    var files = Directory.GetFiles(srcDir)
                                .Where(f => allow.Contains(Path.GetExtension(f) ?? ""))
                                .OrderBy(f => Path.GetFileName(f), new NaturalStringComparer())
                                .ToList();

                    int index = 0;
                    foreach (var row in grp)
                    {
                        int need = row.Count > 0 ? row.Count : 0;           // C<=0 → 不取
                        int take = Math.Min(need, files.Count - index);
                        total += take;
                        index += take;
                    }
                }
            }
            catch { }

            if (total <= 0) { LogEx("本次不提取图片（C 列均为 0 或空），将仅创建目标文件夹。"); }

            BeginWork();

            Task.Run(() =>
            {
                int done = 0, ok = 0, fail = 0, miss = 0, lack = 0;

                foreach (var grp in plan.GroupBy(p => p.Src, StringComparer.OrdinalIgnoreCase))
                {
                    if (_cts?.IsCancellationRequested == true) { LogEx("已停止。"); break; }

                    var srcDir = Path.Combine(input, grp.Key);
                    if (!Directory.Exists(srcDir)) { LogEx($"源目录不存在：{srcDir}"); miss++; continue; }

                    var files = Directory.GetFiles(srcDir)
                             .Where(f => allow.Contains(Path.GetExtension(f) ?? ""))
                             .OrderBy(f => Path.GetFileName(f), new NaturalStringComparer())
                             .ToList();

                    int index = 0; // 连续消耗；若想每行从头，可把 index 重置放到 foreach(row) 内
                    foreach (var row in grp)
                    {
                        if (_cts?.IsCancellationRequested == true) { LogEx("已停止。"); break; }

                        var destRoot = useSourceAsOutput ? srcDir : outputRoot;   // 留空→放源目录；否则放固定输出目录
                        var destDir = Path.Combine(destRoot, row.Dest);
                        try { Directory.CreateDirectory(destDir); } catch { }

                        int need = row.Count > 0 ? row.Count : 0;           // C<=0 → 只建文件夹
                        int take = Math.Min(need, files.Count - index);
                        if (row.Count > 0 && take < need)
                        {
                            LogEx($"警告：{grp.Key} → {row.Dest} 需要 {need}，仅剩 {take}。");
                            lack++;
                        }

                        for (int i = 0; i < take; i++)
                        {
                            if (_cts?.IsCancellationRequested == true) { LogEx("已停止。"); break; }

                            var src = files[index + i];
                            var target = EnsureUnique(Path.Combine(destDir, Path.GetFileName(src)));

                            string current = $"{Path.GetFileName(src)} → {row.Dest}";
                            try
                            {
                                if (rbCopy.Checked) File.Copy(src, target);
                                else File.Move(src, target);
                                ok++;
                            }
                            catch (Exception ex)
                            {
                                LogEx($"失败：{src} → {target}，{ex.Message}");
                                fail++;
                            }
                            done++;
                            UpdateProgressSafe(done, total, current, ok, fail, lack, miss);
                        }

                        if (row.Count <= 0)
                        {
                            LogEx($"创建空文件夹：{grp.Key} → {row.Dest}");
                        }

                        index += take; // C<=0 时 take=0，不前进
                        if (_cts?.IsCancellationRequested == true) break;
                    }

                    var left = files.Count - index;
                    if (left > 0) LogEx($"提示：原文件夹 {grp.Key} 还有剩余 {left} 个未分配。");
                }

                LogEx($"提取完成：总处理 {total}，成功 {ok}，失败 {fail}，未找到源文件夹 {miss}，数量不足 {lack}");
            })
            .ContinueWith(_ => { try { EndWork(); } catch { } }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private static string EnsureUnique(string path)
        {
            if (!File.Exists(path)) return path;
            var dir = Path.GetDirectoryName(path);
            var name = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            int i = 1;
            string candidate;
            do
            {
                candidate = Path.Combine(dir, $"{name}_{i}{ext}");
                i++;
            } while (File.Exists(candidate));
            return candidate;
        }

        // 文件名“自然排序”
        class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string a, string b)
            {
                int ia = 0, ib = 0;
                while (ia < a.Length && ib < b.Length)
                {
                    if (char.IsDigit(a[ia]) && char.IsDigit(b[ib]))
                    {
                        long na = 0, nb = 0;
                        while (ia < a.Length && char.IsDigit(a[ia])) { na = na * 10 + (a[ia] - '0'); ia++; }
                        while (ib < b.Length && char.IsDigit(b[ib])) { nb = nb * 10 + (b[ib] - '0'); ib++; }
                        int cmp = na.CompareTo(nb);
                        if (cmp != 0) return cmp;
                    }
                    else
                    {
                        int cmp = char.ToUpperInvariant(a[ia]).CompareTo(char.ToUpperInvariant(b[ib]));
                        if (cmp != 0) return cmp;
                        ia++; ib++;
                    }
                }
                return (a.Length - ia).CompareTo(b.Length - ib);
            }
        }
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MultiToolWin.Models;
using MultiToolWin.OCR;
using MultiToolWin.PDF;
using MultiToolWin.Services;

namespace MultiToolWin.Pages
{
    public sealed class PagePdf : UserControl
    {
        private readonly Action<string> globalLog;
        private readonly PdfFolderScanner folderScanner = new PdfFolderScanner();
        private readonly PdfGenerationService generationService = new PdfGenerationService();

        private TextBox txtInput;
        private TextBox txtOutput;
        private TextBox txtLog;
        private RadioButton rbImagePdf;
        private RadioButton rbOcrPdf;
        private Button btnBrowseInput;
        private Button btnBrowseOutput;
        private Button btnOpenOutput;
        private Button btnStart;
        private Button btnStop;
        private Button btnCopyLog;
        private Button btnClearLog;
        private ProgressBar progressBar;
        private Label lblCurrent;
        private Label lblProgress;
        private CancellationTokenSource cancellation;

        public PagePdf(Action<string> logger)
        {
            globalLog = logger;
            BuildUi();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { cancellation?.Cancel(); }
                catch (ObjectDisposedException) { }
            }
            base.Dispose(disposing);
        }

        private void BuildUi()
        {
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);

            var rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(12, 8, 12, 8),
                AutoScroll = true
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            rootLayout.Controls.Add(CreateFilesGroup(), 0, 0);
            rootLayout.Controls.Add(CreateOutputTypeGroup(), 0, 1);
            rootLayout.Controls.Add(CreateProgressGroup(), 0, 2);
            rootLayout.Controls.Add(CreateLogGroup(), 0, 3);
            rootLayout.Controls.Add(CreateButtonPanel(), 0, 4);
            Controls.Add(rootLayout);
        }

        private Control CreateFilesGroup()
        {
            const int labelWidth = 100;
            const int buttonWidth = 88;
            var group = new GroupBox
            {
                Text = "文件与目录",
                Dock = DockStyle.Top,
                Height = 104,
                MinimumSize = new System.Drawing.Size(0, 104),
                Padding = new Padding(8),
                Margin = new Padding(0, 0, 0, 8)
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, labelWidth));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, buttonWidth));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, buttonWidth));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            txtInput = CreatePathTextBox();
            txtOutput = CreatePathTextBox();
            btnBrowseInput = CreateButton("浏览…", buttonWidth);
            btnBrowseOutput = CreateButton("浏览…", buttonWidth);
            btnOpenOutput = CreateButton("打开…", buttonWidth);

            btnBrowseInput.Click += (sender, args) => PickFolder(txtInput, "请选择图片根目录");
            btnBrowseOutput.Click += (sender, args) => PickFolder(txtOutput, "请选择 PDF 输出目录");
            btnOpenOutput.Click += (sender, args) => OpenOutputDirectory();

            layout.Controls.Add(CreateLabel("图片根目录："), 0, 0);
            layout.Controls.Add(txtInput, 1, 0);
            layout.Controls.Add(btnBrowseInput, 2, 0);
            layout.Controls.Add(new Panel(), 3, 0);
            layout.Controls.Add(CreateLabel("PDF输出目录："), 0, 1);
            layout.Controls.Add(txtOutput, 1, 1);
            layout.Controls.Add(btnBrowseOutput, 2, 1);
            layout.Controls.Add(btnOpenOutput, 3, 1);
            group.Controls.Add(layout);
            return group;
        }

        private Control CreateOutputTypeGroup()
        {
            var group = new GroupBox
            {
                Text = "输出类型",
                Dock = DockStyle.Top,
                Height = 62,
                Padding = new Padding(10),
                Margin = new Padding(0, 0, 0, 8)
            };
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(4, 5, 0, 0)
            };
            rbImagePdf = new RadioButton
            {
                Text = "普通PDF（仅图片）",
                AutoSize = true,
                Margin = new Padding(0, 0, 24, 0)
            };
            rbOcrPdf = new RadioButton
            {
                Text = "OCR双层PDF（可搜索）",
                AutoSize = true,
                Checked = true
            };
            panel.Controls.Add(rbImagePdf);
            panel.Controls.Add(rbOcrPdf);
            group.Controls.Add(panel);
            return group;
        }

        private Control CreateProgressGroup()
        {
            var group = new GroupBox
            {
                Text = "执行进度",
                Dock = DockStyle.Top,
                Height = 94,
                Padding = new Padding(8),
                Margin = new Padding(0, 0, 0, 8)
            };
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.Controls.Add(new Label
            {
                Text = "当前：",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            }, 0, 0);
            lblCurrent = new Label { Text = "-", AutoSize = true, Anchor = AnchorStyles.Left };
            layout.Controls.Add(lblCurrent, 1, 0);
            lblProgress = new Label
            {
                Text = "进度：0/0",
                AutoSize = true,
                Anchor = AnchorStyles.Left
            };
            layout.Controls.Add(lblProgress, 0, 1);
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 1,
                Value = 0,
                Margin = new Padding(0, 4, 0, 4)
            };
            layout.Controls.Add(progressBar, 1, 1);
            group.Controls.Add(layout);
            return group;
        }

        private Control CreateLogGroup()
        {
            var group = new GroupBox
            {
                Text = "运行日志",
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
                Margin = new Padding(0, 0, 0, 8),
                MinimumSize = new System.Drawing.Size(0, 90)
            };
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0, 3, 0, 0)
            };
            btnClearLog = new Button
            {
                Text = "清空",
                Width = 80,
                Height = 26,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(3, 0, 0, 0)
            };
            btnCopyLog = new Button
            {
                Text = "复制",
                Width = 80,
                Height = 26,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(3, 0, 3, 0)
            };
            btnClearLog.Click += (sender, args) => txtLog.Clear();
            btnCopyLog.Click += (sender, args) =>
            {
                if (!string.IsNullOrEmpty(txtLog.Text))
                    Clipboard.SetText(txtLog.Text);
            };
            buttons.Controls.Add(btnClearLog);
            buttons.Controls.Add(btnCopyLog);
            layout.Controls.Add(txtLog, 0, 0);
            layout.Controls.Add(buttons, 0, 1);
            group.Controls.Add(layout);
            return group;
        }

        private Control CreateButtonPanel()
        {
            var panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 2, 0, 0)
            };
            btnStart = new Button
            {
                Text = "开始生成",
                Width = 100,
                Height = 28,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(8, 3, 0, 3)
            };
            btnStop = new Button
            {
                Text = "停止",
                Width = 88,
                Height = 28,
                FlatStyle = FlatStyle.System,
                Enabled = false,
                Margin = new Padding(0, 3, 0, 3)
            };
            btnStart.Click += BtnStart_Click;
            btnStop.Click += (sender, args) => cancellation?.Cancel();
            panel.Controls.Add(btnStart);
            panel.Controls.Add(btnStop);
            return panel;
        }

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            string inputPath = (txtInput.Text ?? string.Empty).Trim();
            string outputPath = (txtOutput.Text ?? string.Empty).Trim();
            if (!Directory.Exists(inputPath))
            {
                ShowInformation("请选择有效的图片根目录。");
                return;
            }
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                ShowInformation("请选择 PDF 输出目录。");
                return;
            }

            bool includeTextLayer = rbOcrPdf.Checked;
            if (includeTextLayer && !ValidateOcrRequirements()) return;

            cancellation = new CancellationTokenSource();
            SetRunning(true);
            progressBar.Minimum = 0;
            progressBar.Maximum = 1;
            progressBar.Value = 0;
            lblProgress.Text = "进度：0/0";
            lblCurrent.Text = "正在扫描目录…";

            try
            {
                PdfScanResult scanResult = await Task.Run(
                    () => folderScanner.Scan(inputPath),
                    cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();

                if (scanResult.Folders.Count == 0)
                {
                    ShowInformation("未在所选目录的一层或两层子目录中找到支持的图片。");
                    return;
                }

                progressBar.Maximum = Math.Max(1, scanResult.Folders.Count);
                lblProgress.Text = "进度：0/" + scanResult.Folders.Count;
                lblCurrent.Text = "-";

                string layoutName = scanResult.Layout == PdfDirectoryLayout.VolumesAndCases
                    ? "卷→件→图片" : "件→图片";
                LogMessage(string.Format(
                    "开始生成：{0}，识别到 {1} 个件，输出类型：{2}。",
                    layoutName,
                    scanResult.Folders.Count,
                    includeTextLayer ? "OCR双层PDF" : "普通PDF"));

                var progress = new Progress<PdfGenerationProgress>(UpdateProgress);
                PdfGenerationResult result = await generationService.GenerateAsync(
                    scanResult,
                    outputPath,
                    includeTextLayer,
                    progress,
                    cancellation.Token);

                if (result.IsCancelled)
                {
                    string cancelledMessage = string.Format(
                        "任务已取消：生成成功：{0}，同名冲突跳过：{1}，目标已存在跳过：{2}，失败：{3}。",
                        result.GeneratedCount,
                        result.ConflictSkippedCount,
                        result.ExistingSkippedCount,
                        result.FailedCount);
                    LogMessage(cancelledMessage);
                    lblCurrent.Text = "任务已取消";
                    return;
                }

                string message = string.Format(
                    "完成：生成成功：{0}，同名冲突跳过：{1}，目标已存在跳过：{2}，失败：{3}。",
                    result.GeneratedCount,
                    result.ConflictSkippedCount,
                    result.ExistingSkippedCount,
                    result.FailedCount);
                LogMessage(message);
                MessageBox.Show(message, "PDF生成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogMessage("任务已取消。");
                lblCurrent.Text = "任务已取消";
            }
            catch (Exception ex)
            {
                LogMessage("PDF 生成失败：" + ex.Message);
                MessageBox.Show(ex.Message, "PDF生成失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cancellation.Dispose();
                cancellation = null;
                SetRunning(false);
            }
        }

        private bool ValidateOcrRequirements()
        {
            if (!Environment.Is64BitOperatingSystem)
            {
                const string message = "OCR双层 PDF 需要64位 Windows。\r\n普通 PDF 仍可使用。";
                LogMessage(message.Replace("\r\n", " "));
                MessageBox.Show(message, "OCR不可用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            OcrResourceValidationResult validation = OcrResourceValidator.Validate();
            if (!validation.IsValid)
            {
                LogMessage(OcrResourceValidator.UserMessage.Replace(Environment.NewLine, " "));
                foreach (string missingFile in validation.MissingFiles)
                    LogMessage("缺少 OCR 文件：" + missingFile);
                MessageBox.Show(
                    OcrResourceValidator.UserMessage,
                    "OCR组件不完整",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            if (!SearchablePdfBuilder.HasUsableOcrFont())
            {
                const string message =
                    "未找到可用于 OCR 文字层的中文字体（微软雅黑、宋体或黑体）。\r\n普通 PDF 仍可正常生成。";
                LogMessage(message.Replace("\r\n", " "));
                MessageBox.Show(message, "OCR字体不可用", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void UpdateProgress(PdfGenerationProgress info)
        {
            if (!string.IsNullOrWhiteSpace(info.LogMessage))
            {
                LogMessage(info.LogMessage);
                return;
            }

            int value = Math.Max(progressBar.Minimum,
                Math.Min(progressBar.Maximum, info.CompletedFolders));
            progressBar.Value = value;
            lblProgress.Text = string.Format(
                "进度：{0}/{1}", info.CompletedFolders, info.TotalFolders);
            lblCurrent.Text = info.TotalImages > 0
                ? string.Format(
                    "{0}：{1}（{2}/{3}）",
                    info.Message,
                    info.CurrentFolderName,
                    info.CurrentImage,
                    info.TotalImages)
                : string.Format("{0}：{1}", info.Message, info.CurrentFolderName);
            if (info.Message == "已生成")
                LogMessage("已生成：" + info.CurrentFolderName + ".pdf");
        }

        private void SetRunning(bool running)
        {
            btnStart.Enabled = !running;
            btnStop.Enabled = running;
            btnBrowseInput.Enabled = !running;
            btnBrowseOutput.Enabled = !running;
            rbImagePdf.Enabled = !running;
            rbOcrPdf.Enabled = !running;
        }

        private void LogMessage(string message)
        {
            string line = string.Format("[{0:HH:mm:ss}] {1}", DateTime.Now, message);
            try { txtLog.AppendText(line + Environment.NewLine); } catch { }
            try { globalLog?.Invoke("[PDF生成] " + message); } catch { }
        }

        private void ShowInformation(string message)
        {
            LogMessage(message);
            MessageBox.Show(message, "PDF生成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                Margin = new Padding(0, 6, 8, 6)
            };
        }

        private static TextBox CreatePathTextBox()
        {
            return new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 6, 3)
            };
        }

        private static Button CreateButton(string text, int width)
        {
            return new Button
            {
                Text = text,
                Width = width,
                Height = 26,
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(0, 3, 6, 3)
            };
        }

        private static void PickFolder(TextBox target, string description)
        {
            using (var dialog = new FolderBrowserDialog
            {
                Description = description,
                ShowNewFolderButton = true
            })
            {
                if (dialog.ShowDialog() == DialogResult.OK) target.Text = dialog.SelectedPath;
            }
        }

        private void OpenOutputDirectory()
        {
            string path = (txtOutput.Text ?? string.Empty).Trim();
            if (!Directory.Exists(path))
            {
                ShowInformation("请先选择已存在的 PDF 输出目录。");
                return;
            }

            try { Process.Start("explorer.exe", path); }
            catch (Exception ex) { ShowInformation("打开目录失败：" + ex.Message); }
        }
    }
}

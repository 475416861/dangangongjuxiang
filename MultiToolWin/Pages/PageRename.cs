using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static MultiToolWin.Utils.ExcelUtils;

namespace MultiToolWin.Pages
{
    public class PageRename : UserControl
    {
        private Action<string> Log;

        private TextBox txtRoot;
        private Button btnPickRoot;

        // 功能A
        private Button btnExportExcel, btnApplyExcel;
        private RadioButton rdoFolderMode, rdoFileMode, rdoTwoLevelFolderMode;

        // 功能B
        private TextBox txtSep;
        private NumericUpDown numDigits, numStart;
        private Button btnRenameFiles, btnPreview;
        private RadioButton rdoLevel1, rdoLevel2;
        private RadioButton rdoAllTypes, rdoSelectTypes;
        private CheckBox chkJpg, chkTif, chkPdf, chkPng;
        private ProgressBar progressBar;
        private Label lblProgress;
        private TextBox txtLog;
        private Button btnStop;
        private bool stopRequested = false;
        private string lastExportPath = null;
        // 日志保存：内存里保留全量，UI只显示最新10条
        private List<string> allLogs = new List<string>();
        private int maxUiLogLines = 10;

        private sealed class RenameMapRow
        {
            public string OldName { get; set; }
            public string NewName { get; set; }
            public string OriginalRelativePath { get; set; }
        }

        private sealed class TwoLevelFolderItem
        {
            public string Name { get; set; }
            public string RelativePath { get; set; }
        }



        public PageRename(Action<string> log)
        {
            this.Log = log;
            BuildUI();
        }
        private void ExportListToExcel(List<string> items, string filePath, string sheetName)
        {
            using (var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet(sheetName);

                // 表头
                var header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("旧名称");
                header.CreateCell(1).SetCellValue("新名称");

                // 内容
                for (int i = 0; i < items.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(items[i]);
                    row.CreateCell(1).SetCellValue(""); // 新名称列留空，供人工填写
                }

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
        }

        private void ExportTwoLevelFolderListToExcel(string root, string filePath)
        {
            var comparer = new NaturalStringComparer();
            var items = new List<TwoLevelFolderItem>();
            var volumes = Directory.GetDirectories(root)
                .OrderBy(path => Path.GetFileName(path), comparer)
                .ToArray();

            foreach (var volume in volumes)
            {
                var volumeName = Path.GetFileName(volume);
                items.Add(new TwoLevelFolderItem
                {
                    Name = volumeName,
                    RelativePath = volumeName
                });

                var folders = Directory.GetDirectories(volume)
                    .OrderBy(path => Path.GetFileName(path), comparer);
                foreach (var folder in folders)
                {
                    items.Add(new TwoLevelFolderItem
                    {
                        Name = Path.GetFileName(folder),
                        RelativePath = Path.Combine(volumeName, Path.GetFileName(folder))
                    });
                }
            }

            using (var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("两层文件夹清单");
                var header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("旧名称");
                header.CreateCell(1).SetCellValue("新名称");
                header.CreateCell(2).SetCellValue("原始相对路径");

                for (int i = 0; i < items.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(items[i].Name);
                    row.CreateCell(1).SetCellValue(string.Empty);
                    row.CreateCell(2).SetCellValue(items[i].RelativePath);
                }

                sheet.SetColumnHidden(2, true);
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }

            LogEx($"导出两层文件夹清单成功，共 {volumes.Length} 个卷目录、{items.Count - volumes.Length} 个件目录。");
        }

        private Dictionary<string, string> LoadMapFromExcel(string filePath)
        {
            var map = new Dictionary<string, string>();
            using (var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(filePath))
            {
                var sheet = workbook.GetSheetAt(0);

                // 从第1行开始读（跳过表头）
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;

                    var oldCell = row.GetCell(0);
                    var newCell = row.GetCell(1);

                    if (oldCell == null) continue;

                    var oldName = oldCell.ToString().Trim();
                    var newName = newCell?.ToString().Trim() ?? "";

                    if (!string.IsNullOrEmpty(oldName))
                    {
                        map[oldName] = newName;
                    }
                }
            }

            return map;
        }

        private List<RenameMapRow> LoadTwoLevelRenameRows(string filePath, out bool hasOriginalRelativePathColumn)
        {
            var rows = new List<RenameMapRow>();
            hasOriginalRelativePathColumn = false;

            using (var stream = File.OpenRead(filePath))
            using (var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(stream))
            {
                var sheet = workbook.GetSheetAt(0);
                var header = sheet?.GetRow(sheet.FirstRowNum);
                if (header == null) return rows;

                int pathColumn = -1;
                for (int column = header.FirstCellNum; column < header.LastCellNum; column++)
                {
                    var headerValue = header.GetCell(column)?.ToString()?.Trim();
                    if (string.Equals(headerValue, "原始相对路径", StringComparison.OrdinalIgnoreCase))
                    {
                        pathColumn = column;
                        hasOriginalRelativePathColumn = true;
                        break;
                    }
                }

                if (!hasOriginalRelativePathColumn) return rows;

                for (int rowIndex = sheet.FirstRowNum + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    var oldName = row.GetCell(0)?.ToString()?.Trim() ?? string.Empty;
                    var newName = row.GetCell(1)?.ToString()?.Trim() ?? string.Empty;
                    var relativePath = row.GetCell(pathColumn)?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrEmpty(oldName) && string.IsNullOrEmpty(relativePath)) continue;

                    rows.Add(new RenameMapRow
                    {
                        OldName = oldName,
                        NewName = newName,
                        OriginalRelativePath = relativePath
                    });
                }
            }

            return rows;
        }
        private List<FileInfo> CollectTargetFiles(string root)
        {
            var result = new List<FileInfo>();

            if (!Directory.Exists(root))
            {
                LogEx("根目录不存在。");
                return result;
            }

            // 1. 确定扩展名范围
            HashSet<string> allowExt;
            if (rdoAllTypes.Checked)
            {
                allowExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg", ".jpeg", ".tif", ".tiff", ".pdf", ".png"
        };
            }
            else
            {
                allowExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (chkJpg.Checked) allowExt.Add(".jpg");
                if (chkJpg.Checked) allowExt.Add(".jpeg"); // 兼容
                if (chkTif.Checked) { allowExt.Add(".tif"); allowExt.Add(".tiff"); }
                if (chkPdf.Checked) allowExt.Add(".pdf");
                if (chkPng.Checked) allowExt.Add(".png");
            }

            // 2. 根据层级模式收集文件
            if (rdoLevel1.Checked)
            {
                // 一级子文件夹
                foreach (var dir in Directory.GetDirectories(root))
                {
                    foreach (var file in Directory.GetFiles(dir))
                    {
                        var ext = Path.GetExtension(file);
                        if (allowExt.Contains(ext))
                            result.Add(new FileInfo(file));
                    }
                }
            }
            else if (rdoLevel2.Checked)
            {
                // 二级子文件夹
                foreach (var dir in Directory.GetDirectories(root))
                {
                    foreach (var sub in Directory.GetDirectories(dir))
                    {
                        foreach (var file in Directory.GetFiles(sub))
                        {
                            var ext = Path.GetExtension(file);
                            if (allowExt.Contains(ext))
                                result.Add(new FileInfo(file));
                        }
                    }
                }
            }

            return result;
        }
        private void RenameFiles(string root, bool preview)
        {
            var files = CollectTargetFiles(root);
            if (files.Count == 0)
            {
                LogEx("未找到符合条件的文件。");
                return;
            }

            progressBar.Value = 0;
            progressBar.Maximum = files.Count;
            int success = 0, skip = 0, fail = 0;

            string sep = txtSep.Text;
            int digits = (int)numDigits.Value;
            int start = (int)numStart.Value;

            var groups = files.GroupBy(f => f.Directory.FullName);

            int previewCount = 0; // 预览条数计数

            foreach (var group in groups)
            {
                string folderName = Path.GetFileName(group.Key);
                int counter = start;

                var ordered = group.OrderBy(f => f.Name, new NaturalStringComparer());

                foreach (var fi in ordered)
                {
                    try
                    {
                        string newName = folderName + sep + counter.ToString().PadLeft(digits, '0') + fi.Extension.ToLower();
                        string newPath = Path.Combine(fi.DirectoryName, newName);

                        if (preview)
                        {
                            if (previewCount < 10)  // ✅ 仅显示前 10 条
                            {
                                LogEx($"预览：{fi.Name} → {newName}");
                                previewCount++;
                            }
                        }
                        else
                        {
                            if (File.Exists(newPath))
                            {
                                skip++;
                                LogEx($"跳过：目标已存在 {newName}");
                            }
                            else
                            {
                                File.Move(fi.FullName, newPath);
                                success++;
                            }
                        }

                        counter++;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        LogEx($"失败：{fi.Name}，原因：{ex.Message}");
                    }

                    progressBar.Value++;
                    if (progressBar.Value % 50 == 0 || progressBar.Value == progressBar.Maximum)
                    {
                        lblProgress.Text = $"进度：{progressBar.Value}/{progressBar.Maximum}";
                        Application.DoEvents();
                    }

                    if (stopRequested)
                    {
                        LogEx("用户停止操作。");
                        return;
                    }
                }
            }

            if (!preview)
            {
                LogEx($"完成：成功 {success}，跳过 {skip}，失败 {fail}");
            }
            else
            {
                if (files.Count > 10)
                {
                    LogEx($"预览完成，仅显示前 10 条，共 {files.Count} 个文件。");
                }
                else
                {
                    LogEx("预览完成，未修改任何文件。");
                }
            }
        }



        private void BuildUI()
        {
            this.Dock = DockStyle.Fill;
            this.Font = new Font("Microsoft YaHei UI", 9F);
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(8),
                AutoScroll = true
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // 根目录实际首选高度为 64px；使用默认 100px 会挤占执行进度区域。
            var rootGroup = new GroupBox { Text = "工作目录", Dock = DockStyle.Fill, Height = 64, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 4) };
            var rootLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
            rootLayout.Controls.Add(new Label { Text = "根目录：", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0, 4, 8, 4) }, 0, 0);
            txtRoot = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 8, 4) };
            btnPickRoot = new Button { Text = "浏览…", Width = 80, Height = 26, FlatStyle = FlatStyle.System, Margin = new Padding(0, 3, 0, 3) };
            btnPickRoot.Click += BtnPickRoot_Click;
            rootLayout.Controls.Add(txtRoot, 1, 0);
            rootLayout.Controls.Add(btnPickRoot, 2, 0);
            rootGroup.Controls.Add(rootLayout);
            layout.Controls.Add(rootGroup, 0, 0);

            var groupA = new GroupBox { Text = "功能A：Excel 映射", Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
            var layoutA = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
            rdoFolderMode = new RadioButton { Text = "导出文件夹清单", Checked = true, AutoSize = true, Margin = new Padding(0, 3, 12, 3) };
            rdoFileMode = new RadioButton { Text = "导出文件清单", AutoSize = true, Margin = new Padding(0, 3, 10, 3) };
            rdoTwoLevelFolderMode = new RadioButton { Text = "两层文件夹清单", AutoSize = true, Margin = new Padding(0, 3, 10, 3) };
            btnExportExcel = new Button { Text = "导出Excel", Width = 100, Height = 28, FlatStyle = FlatStyle.System };
            btnApplyExcel = new Button { Text = "应用Excel映射", Width = 120, Height = 28, FlatStyle = FlatStyle.System };
            btnExportExcel.Click += BtnExportExcel_Click;
            btnApplyExcel.Click += BtnApplyExcel_Click;
            var modesPanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false };
            modesPanel.Controls.Add(rdoFolderMode);
            modesPanel.Controls.Add(rdoFileMode);
            modesPanel.Controls.Add(rdoTwoLevelFolderMode);
            var mappingButtons = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 6, 0, 0) };
            mappingButtons.Controls.Add(btnApplyExcel);
            mappingButtons.Controls.Add(btnExportExcel);
            layoutA.Controls.Add(modesPanel, 0, 0);
            layoutA.Controls.Add(mappingButtons, 0, 1);
            groupA.Controls.Add(layoutA);
            layout.Controls.Add(groupA, 0, 1);

            var groupB = new GroupBox { Text = "功能B：规则改名", Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8) };
            var layoutB = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };

            // 层级模式
            var lblLevelMode = new Label
            {
                Text = "层级模式：",
                Width = 80,                       // 固定宽度，保证和“连接符”对齐
                TextAlign = ContentAlignment.MiddleRight
            };
            rdoLevel1 = new RadioButton { Text = "一级子文件夹", Checked = true, AutoSize = true };
            rdoLevel2 = new RadioButton { Text = "二级子文件夹", AutoSize = true };

            var panelLevel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top
            };
            panelLevel.Controls.Add(lblLevelMode);
            panelLevel.Controls.Add(rdoLevel1);
            panelLevel.Controls.Add(rdoLevel2);
            layoutB.Controls.Add(panelLevel, 0, 0);

            // 连接符 + 位数 + 起始号
            var lblSep = new Label
            {
                Text = "连接符：",
                Width = 80,
                Height = 25,                       // 和输入框统一高度
                AutoSize = false,                   // 必须关闭，否则 Height 无效
                TextAlign = ContentAlignment.MiddleRight
            };
            txtSep = new TextBox { Width = 50, Text = "-", Height = 25 };

            var lblDigits = new Label
            {
                Text = "序号位数：",
                Width = 70,
                Height = 25,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleRight
            };
            numDigits = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 10,
                Value = 3,
                Width = 60,
                Height = 25
            };

            var lblStart = new Label
            {
                Text = "起始号：",
                Width = 70,
                Height = 25,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleRight
            };
            numStart = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 9999,
                Value = 1,
                Width = 60,
                Height = 25
            };

            var panelSep = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top
            };
            panelSep.Controls.Add(lblSep);
            panelSep.Controls.Add(txtSep);
            panelSep.Controls.Add(lblDigits);
            panelSep.Controls.Add(numDigits);
            panelSep.Controls.Add(lblStart);
            panelSep.Controls.Add(numStart);

            layoutB.Controls.Add(panelSep, 0, 1);



            // 处理范围
            var panelScope = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Dock = DockStyle.Top, Padding = new Padding(0, 4, 0, 0) };
            rdoAllTypes = new RadioButton { Text = "所有支持的文件类型", Checked = true, AutoSize = true };
            rdoSelectTypes = new RadioButton { Text = "仅选中的文件类型：", AutoSize = true };
            var panelTypes = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            chkJpg = new CheckBox { Text = "JPG", Checked = true, AutoSize = true };
            chkTif = new CheckBox { Text = "TIF", Checked = true, AutoSize = true };
            chkPdf = new CheckBox { Text = "PDF", AutoSize = true };
            chkPng = new CheckBox { Text = "PNG", AutoSize = true };
            panelTypes.Controls.Add(chkJpg);
            panelTypes.Controls.Add(chkTif);
            panelTypes.Controls.Add(chkPdf);
            panelTypes.Controls.Add(chkPng);

            panelScope.Controls.Add(rdoAllTypes);
            panelScope.Controls.Add(rdoSelectTypes);
            panelScope.Controls.Add(panelTypes);
            layoutB.Controls.Add(panelScope, 0, 2);

            // 操作按钮
            var panelOps = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Top, Padding = new Padding(0, 6, 0, 0) };
            btnPreview = new Button { Text = "预览", Width = 100, Height = 28, FlatStyle = FlatStyle.System };
            btnRenameFiles = new Button { Text = "开始执行", Width = 100, Height = 28, FlatStyle = FlatStyle.System };
            btnStop = new Button { Text = "停止", Width = 100, Height = 28, FlatStyle = FlatStyle.System };
            btnPreview.Click += (s, e) => DoPreview();
            btnRenameFiles.Click += BtnRenameFiles_Click;
            btnStop.Click += BtnStop_Click;
            panelOps.Controls.Add(btnStop);
            panelOps.Controls.Add(btnRenameFiles);
            panelOps.Controls.Add(btnPreview);
            layoutB.Controls.Add(panelOps, 0, 3);
            groupB.Controls.Add(layoutB);
            layout.Controls.Add(groupB, 0, 2);

            // 标题、内边距和三行固定控件合计至少需要 136px；避免底部操作行被百分比日志行挤压。
            var logGroup = new GroupBox { Text = "执行进度", Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 8), MinimumSize = new Size(0, 136) };
            var logLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
            logLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 24F));
            logLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            logLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            logLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            progressBar = new ProgressBar { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 4) };
            lblProgress = new Label { Text = "进度：0/0", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            txtLog = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Margin = new Padding(0, 0, 0, 4) };
            var panelLogOps = new FlowLayoutPanel { AutoSize = false, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Dock = DockStyle.Fill, Padding = new Padding(0, 4, 0, 0), Margin = new Padding(0) };
            var btnExportLog = new Button { Text = "导出日志", Width = 100, Height = 26, Margin = new Padding(0), FlatStyle = FlatStyle.System };
            btnExportLog.Click += BtnExportLog_Click;
            panelLogOps.Controls.Add(btnExportLog);
            logLayout.Controls.Add(progressBar, 0, 0);
            logLayout.Controls.Add(lblProgress, 0, 1);
            logLayout.Controls.Add(txtLog, 0, 2);
            logLayout.Controls.Add(panelLogOps, 0, 3);
            logGroup.Controls.Add(logLayout);
            layout.Controls.Add(logGroup, 0, 3);
            this.Controls.Add(layout);
        }

        private void BtnPickRoot_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtRoot.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            var root = txtRoot.Text;
            if (!Directory.Exists(root))
            {
                LogEx("请选择有效的根目录。");
                return;
            }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 文件|*.xlsx";
                sfd.FileName = "导出清单.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (rdoTwoLevelFolderMode.Checked)
                        {
                            ExportTwoLevelFolderListToExcel(root, sfd.FileName);
                        }
                        else if (rdoFolderMode.Checked)
                        {
                            var folders = Directory.GetDirectories(root)
                                .OrderBy(path => Path.GetFileName(path), new NaturalStringComparer())
                                .ToArray();
                            ExportListToExcel(folders.Select(f => Path.GetFileName(f)).ToList(), sfd.FileName, "文件夹清单");
                            LogEx($"导出文件夹清单成功，共 {folders.Length} 个文件夹。");
                        }
                        else if (rdoFileMode.Checked)
                        {
                            var files = Directory.GetFiles(root)
                                .OrderBy(path => Path.GetFileName(path), new NaturalStringComparer())
                                .ToArray();
                            ExportListToExcel(files.Select(f => Path.GetFileName(f)).ToList(), sfd.FileName, "文件清单");
                            LogEx($"导出文件清单成功，共 {files.Length} 个文件。");
                        }

                        // ✅ 记住路径
                        lastExportPath = sfd.FileName;
                    }
                    catch (IOException ex)
                    {
                        if (File.Exists(sfd.FileName))
                            LogEx("导出失败：文件可能正在被 Excel/WPS 或其他程序使用，请关闭相关程序后重试。");
                        else
                            LogEx("导出失败（文件读写错误）：" + ex.Message);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        LogEx("导出失败：没有权限访问目标文件，请检查目录权限或文件是否为只读。" + Environment.NewLine + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        LogEx("导出失败：" + ex.Message);
                    }
                }
            }
        }



        private void BtnApplyExcel_Click(object sender, EventArgs e)
        {
            var root = txtRoot.Text;
            if (!Directory.Exists(root))
            {
                LogEx("请选择有效的根目录。");
                return;
            }

            string pathToUse = lastExportPath;

            // 如果没有记录路径，或者文件不存在 → 让用户选
            if (string.IsNullOrEmpty(pathToUse) || !File.Exists(pathToUse))
            {
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Excel 文件|*.xlsx";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        pathToUse = ofd.FileName;
                        lastExportPath = pathToUse; // 顺便更新
                    }
                    else
                    {
                        LogEx("未选择Excel文件，操作取消。");
                        return;
                    }
                }
            }

            try
            {
                bool hasOriginalRelativePathColumn;
                var twoLevelRows = LoadTwoLevelRenameRows(pathToUse, out hasOriginalRelativePathColumn);
                if (hasOriginalRelativePathColumn)
                {
                    ApplyTwoLevelExcelMap(root, twoLevelRows);
                    return;
                }

                var map = LoadMapFromExcel(pathToUse);

                int success = 0, skip = 0, fail = 0;

                // 同时支持文件夹和文件
                var dirs = Directory.GetDirectories(root);
                var files = Directory.GetFiles(root);

                foreach (var kv in map)
                {
                    var oldName = kv.Key;
                    var newName = kv.Value;

                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        skip++;
                        LogEx($"跳过：{oldName} → (新名称为空)");
                        continue;
                    }

                    // 优先文件夹
                    var oldDir = dirs.FirstOrDefault(d => Path.GetFileName(d) == oldName);
                    if (oldDir != null)
                    {
                        var newDir = Path.Combine(root, newName);
                        if (Directory.Exists(newDir))
                        {
                            skip++;
                            LogEx($"跳过：目标已存在 {newDir}");
                            continue;
                        }
                        Directory.Move(oldDir, newDir);
                        success++;
                        LogEx($"文件夹重命名成功：{oldName} → {newName}");
                        continue;
                    }

                    // 再文件
                    var oldFile = files.FirstOrDefault(f => Path.GetFileName(f) == oldName);
                    if (oldFile != null)
                    {
                        var newFile = Path.Combine(root, newName);
                        if (File.Exists(newFile))
                        {
                            skip++;
                            LogEx($"跳过：目标已存在 {newFile}");
                            continue;
                        }
                        File.Move(oldFile, newFile);
                        success++;
                        LogEx($"文件重命名成功：{oldName} → {newName}");
                        continue;
                    }

                    fail++;
                    LogEx($"未找到：{oldName}");
                }

                LogEx($"完成：成功 {success}，跳过 {skip}，失败 {fail}");
            }
            catch (Exception ex)
            {
                LogEx("应用Excel映射失败：" + ex.Message);
            }
        }

        private void ApplyTwoLevelExcelMap(string root, List<RenameMapRow> rows)
        {
            int success = 0, skip = 0, fail = 0;

            foreach (var row in rows.OrderByDescending(item => GetRelativePathDepth(item.OriginalRelativePath)))
            {
                var displayPath = string.IsNullOrWhiteSpace(row.OriginalRelativePath)
                    ? row.OldName
                    : row.OriginalRelativePath;

                if (string.IsNullOrWhiteSpace(row.NewName))
                {
                    skip++;
                    LogEx($"跳过：{displayPath}（未填写新名称）");
                    continue;
                }

                if (!IsSafeFolderName(row.NewName))
                {
                    fail++;
                    LogEx($"失败：{displayPath} → {row.NewName}，新名称必须是有效的单个文件夹名称。");
                    continue;
                }

                string oldDirectory;
                string normalizedRelativePath;
                if (!TryResolveTwoLevelDirectory(root, row.OriginalRelativePath, out oldDirectory, out normalizedRelativePath))
                {
                    fail++;
                    LogEx($"失败：{displayPath} → {row.NewName}，原始相对路径无效或超出根目录。");
                    continue;
                }

                if (!Directory.Exists(oldDirectory))
                {
                    fail++;
                    LogEx($"失败：{normalizedRelativePath} → {row.NewName}，原目录不存在。");
                    continue;
                }

                try
                {
                    var parentDirectory = Directory.GetParent(oldDirectory)?.FullName;
                    var newDirectory = Path.Combine(parentDirectory, row.NewName);
                    if (Directory.Exists(newDirectory) || File.Exists(newDirectory))
                    {
                        skip++;
                        LogEx($"跳过：{normalizedRelativePath} → {row.NewName}，目标已存在。");
                        continue;
                    }

                    Directory.Move(oldDirectory, newDirectory);
                    success++;
                    LogEx($"文件夹重命名成功：{normalizedRelativePath} → {row.NewName}");
                }
                catch (Exception ex)
                {
                    fail++;
                    LogEx($"失败：{normalizedRelativePath} → {row.NewName}，{ex.Message}");
                }
            }

            LogEx($"两层文件夹映射完成：成功 {success}，跳过 {skip}，失败 {fail}");
        }

        private static int GetRelativePathDepth(string relativePath)
        {
            return (relativePath ?? string.Empty)
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Length;
        }

        private static bool IsSafeFolderName(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName)) return false;
            if (folderName == "." || folderName == "..") return false;
            if (Path.IsPathRooted(folderName)) return false;
            if (folderName.IndexOf('\\') >= 0 || folderName.IndexOf('/') >= 0) return false;
            if (folderName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;
            return !folderName.EndsWith(".", StringComparison.Ordinal) && !folderName.EndsWith(" ", StringComparison.Ordinal);
        }

        private static bool TryResolveTwoLevelDirectory(string root, string relativePath, out string directory, out string normalizedRelativePath)
        {
            directory = null;
            normalizedRelativePath = null;

            var input = (relativePath ?? string.Empty).Trim();
            if (input.Length == 0 || Path.IsPathRooted(input)) return false;

            var parts = input.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 2 || parts.Any(part => part == "." || part == "..")) return false;

            try
            {
                var rootPrefix = GetRootPathPrefix(root);
                var fullPath = Path.GetFullPath(Path.Combine(rootPrefix, string.Join(Path.DirectorySeparatorChar.ToString(), parts)));
                if (!fullPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase)) return false;

                directory = fullPath;
                normalizedRelativePath = fullPath.Substring(rootPrefix.Length);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetRootPathPrefix(string root)
        {
            var rootPath = Path.GetFullPath(root);
            return rootPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                || rootPath.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? rootPath
                : rootPath + Path.DirectorySeparatorChar;
        }



        private void BtnRenameFiles_Click(object sender, EventArgs e)
        {
            var root = txtRoot.Text;
            if (!Directory.Exists(root))
            {
                LogEx("请选择有效的根目录。");
                return;
            }

            stopRequested = false;
            RenameFiles(root, false);
        }


        private void BtnStop_Click(object sender, EventArgs e)
        {
            stopRequested = true;
            LogEx("停止任务请求");
        }

        private void DoPreview()
        {
            var root = txtRoot.Text;
            if (!Directory.Exists(root))
            {
                LogEx("请选择有效的根目录。");
                return;
            }

            stopRequested = false;
            RenameFiles(root, true);   // 预览模式
        }

        private void BtnExportLog_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "文本文件|*.txt";
                sfd.FileName = "操作日志.txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        File.WriteAllLines(sfd.FileName, allLogs);
                        MessageBox.Show("日志已导出：" + sfd.FileName, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void LogEx(string msg)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";

            // 内存保存完整日志
            allLogs.Add(line);

            // UI 显示：只保留最新10条
            txtLog.AppendText(line + Environment.NewLine);

            if (txtLog.Lines.Length > maxUiLogLines)
            {
                var latest = txtLog.Lines.Skip(txtLog.Lines.Length - maxUiLogLines).ToArray();
                txtLog.Lines = latest;
                txtLog.SelectionStart = txtLog.Text.Length;
                txtLog.ScrollToCaret();
            }
        }
    }
}



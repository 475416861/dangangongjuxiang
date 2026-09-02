using iTextSharp.text.pdf;
using MultiToolWin.Utils;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MultiToolWin.Pages
{
    public class PageCompare : UserControl
    {
        private Action<string> Log;

        private TextBox txtExcel, txtRoot;
        private CheckedListBox chkFormats;
        private Button btnExcel, btnRoot, btnStart, btnExport;
        private DataGridView grid;
        private TableLayoutPanel _layout;

        // —— 新增：自动识别出来的列名（不再用下拉框）——
        private string _folderColName = null;
        private string _expectedColName = null;

        private sealed class CompareTarget
        {
            public string FullPath { get; set; }
            public string RelativePath { get; set; }
            public string Name { get; set; }
        }

        private sealed class CompareTargetResolution
        {
            public CompareTarget Target { get; set; }
            public string Status { get; set; }
        }

        public PageCompare(Action<string> logger)
        {
            Log = logger;
            BuildUI();
        }

        private void BuildUI()
        {
            _layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                Padding = new Padding(12, 10, 12, 10)
            };
            _layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            _layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // ===== 0 行：Excel 文件 =====
            _layout.Controls.Add(new Label { Text = "Excel 文件：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) }, 0, 0);

            txtExcel = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            UiStyle.StyleBrowseRow(txtExcel, null);
            _layout.Controls.Add(txtExcel, 1, 0);

            btnExcel = new Button { Text = "浏览…" };
            UiStyle.StyleBrowseRow(null, btnExcel);
            btnExcel.Click += (s, e) =>
            {
                var ofd = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtExcel.Text = ofd.FileName;
                    var header = ExcelUtils.GetHeader(ofd.FileName);
                    GuessMapping(header, out _folderColName, out _expectedColName);

                    if (_folderColName == null || _expectedColName == null)
                        Log("未能从表头自动识别到“文件夹列 / 应有数量列”，默认使用前两列。");
                    else
                        Log($"已自动识别：文件夹列=[{_folderColName}]，应有数量列=[{_expectedColName}]。");
                }
            };
            _layout.Controls.Add(btnExcel, 2, 0);

            // 放大两行之间的竖向间距（覆盖 UiStyle 的默认 Margin）
            txtExcel.Margin = new Padding(0, 6, 6, 12);
            btnExcel.Margin = new Padding(0, 6, 0, 12);

            // ===== 1 行：根目录 =====
            _layout.Controls.Add(new Label { Text = "根目录：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) }, 0, 1);

            txtRoot = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            UiStyle.StyleBrowseRow(txtRoot, null);
            _layout.Controls.Add(txtRoot, 1, 1);

            btnRoot = new Button { Text = "浏览…" };
            UiStyle.StyleBrowseRow(null, btnRoot);
            btnRoot.Click += (s, e) =>
            {
                using (var fbd = new FolderBrowserDialog())
                    if (fbd.ShowDialog() == DialogResult.OK) txtRoot.Text = fbd.SelectedPath;
            };
            _layout.Controls.Add(btnRoot, 2, 1);

            // 再加点与下一块的间距
            txtRoot.Margin = new Padding(0, 6, 6, 18);
            btnRoot.Margin = new Padding(0, 6, 0, 18);

            // ===== 2 行：格式统计 =====
            _layout.Controls.Add(new Label { Text = "格式统计：", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 0) }, 0, 2);

            chkFormats = new CheckedListBox
            {
                Height = 96,
                Margin = new Padding(0, 0, 0, 0),
                Anchor = AnchorStyles.Left | AnchorStyles.Top
            };
            chkFormats.Items.AddRange(new object[] { "JPG", "PNG", "TIF", "GIF", "PDF" });
            for (int i = 0; i < chkFormats.Items.Count; i++) chkFormats.SetItemChecked(i, true);
            UiStyle.StyleCheckedList(chkFormats, height: 96);
            _layout.Controls.Add(chkFormats, 1, 2);

            // ===== 3 行：按钮条（新行，表格上方）=====
            btnStart = new Button
            {
                Text = "开始对比",
                Width = 90,
                Height = 28,
                Margin = new Padding(0, 0, 8, 0),
                FlatStyle = FlatStyle.System
            };
            btnStart.Click += (s, e) => DoCompare();

            btnExport = new Button
            {
                Text = "导出",
                Width = 90,
                Height = 28,
                Margin = new Padding(0, 0, 0, 0),
                FlatStyle = FlatStyle.System
            };
            btnExport.Click += btnExport_Click;

            var buttonsBar = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 6),
                Margin = new Padding(0, 6, 0, 0)
            };
            buttonsBar.Controls.Add(btnStart);
            buttonsBar.Controls.Add(btnExport);

            // 放到第3行第3列（索引2），自动靠右
            buttonsBar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            buttonsBar.Margin = new Padding(0, 6, 2, 6);
            _layout.Controls.Add(buttonsBar, 2, 3);   // ✅ 右侧
         


            // ===== 4 行：结果表 =====
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            grid.Columns.Add("Folder", "文件夹");
            grid.Columns.Add("Expected", "应有");
            grid.Columns.Add("Actual", "实际");
            grid.Columns.Add("Diff", "差值");
            grid.Columns.Add("Match", "匹配");

            UiStyle.StyleGrid(grid);
            grid.RowTemplate.Height = 26;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 248, 250);

            grid.Columns[0].FillWeight = 40;
            grid.Columns[1].FillWeight = 15;
            grid.Columns[2].FillWeight = 15;
            grid.Columns[3].FillWeight = 15;
            grid.Columns[4].FillWeight = 15;

            UiStyle.AlignRight(grid.Columns[1], grid.Columns[2], grid.Columns[3]);
            UiStyle.AlignCenter(grid.Columns[4]);

            _layout.SetColumnSpan(grid, 3);
            _layout.Controls.Add(grid, 0, 4);

            // 行样式
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0 Excel
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 1 根目录
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 2 格式统计
            _layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 3 按钮条
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // 4 表格

            Controls.Add(_layout);
        }


        // —— 自动猜测映射列 —— 
        private static void GuessMapping(List<string> header, out string folderCol, out string expectedCol)
        {
            folderCol = null;
            expectedCol = null;
            if (header == null || header.Count == 0) return;

            // 先做小写副本便于匹配
            var lower = header.Select(h => (orig: h, low: (h ?? "").ToLowerInvariant())).ToList();

            // 关键词集合（尽量覆盖常见中文/英文）
            string[] folderKeys = { "文件夹", "目录", "文件", "名称", "名字", "档名", "folder", "dir", "name", "path" };
            string[] expectedKeys = { "应有", "应", "页", "页数", "数量", "数", "expected", "count", "pages", "page", "qty", "quantity" };

            // 规则：出现任何一个关键词就算匹配。若多列匹配，取第一命中。
            foreach (var kv in lower)
                if (folderCol == null && folderKeys.Any(k => kv.low.Contains(k)))
                    folderCol = kv.orig;

            foreach (var kv in lower)
                if (expectedCol == null && expectedKeys.Any(k => kv.low.Contains(k)))
                    expectedCol = kv.orig;

            // 回退：如果还没识别出来，就用前两列（若存在）
            if (folderCol == null) folderCol = header[0];
            if (expectedCol == null)
            {
                expectedCol = header.Count > 1 ? header[1] : header[0];
            }
        }

        private void DoCompare()
        {
            var excel = txtExcel.Text?.Trim();
            var root = txtRoot.Text?.Trim();

            if (!File.Exists(excel)) { Log("请选择有效的 Excel 文件。"); return; }
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) { Log("请选择有效的根目录。"); return; }

            // 如果用户直接点“开始对比”而未重新选择 Excel（导致列名为空），这里补一次识别
            if (_folderColName == null || _expectedColName == null)
            {
                var hdr = ExcelUtils.GetHeader(excel);
                GuessMapping(hdr, out _folderColName, out _expectedColName);
                Log($"自动识别字段映射：文件夹列 = [{_folderColName}]，应有数量列 = [{_expectedColName}]。");
            }

            var rows = ExcelUtils.ReadExpected(excel, _folderColName, _expectedColName);
            if (rows.Count == 0) { Log("Excel 数据为空或映射列错误。"); return; }

            // 选择的格式
            var useJPG = chkFormats.GetItemChecked(chkFormats.Items.IndexOf("JPG"));
            var usePNG = chkFormats.GetItemChecked(chkFormats.Items.IndexOf("PNG"));
            var useTIF = chkFormats.GetItemChecked(chkFormats.Items.IndexOf("TIF"));
            var useGIF = chkFormats.GetItemChecked(chkFormats.Items.IndexOf("GIF"));
            var usePDF = chkFormats.GetItemChecked(chkFormats.Items.IndexOf("PDF"));

            var targets = CollectCompareTargets(root);
            if (targets == null) return;
            Log($"已在根目录前两层发现 {targets.Count} 个可校对目录。");

            grid.Rows.Clear();
            int match = 0;

            foreach (var (folder, expected) in rows)
            {
                var resolution = ResolveCompareTarget(root, folder, targets);
                if (resolution.Target == null)
                {
                    grid.Rows.Add(folder, expected, "—", "—", resolution.Status);
                    continue;
                }

                var actual = CountActual(resolution.Target.FullPath, useJPG, usePNG, useTIF, useGIF, usePDF);

                int diff = actual - expected;
                bool ok = diff == 0;
                if (ok) match++;

                grid.Rows.Add(folder, expected, actual, diff, ok ? "✔" : "✖");

                if (!ok)
                    Log($"数量不一致：{resolution.Target.RelativePath}，应有：{expected}，实际：{actual}，差值：{diff}");
            }

            Log($"对比完成：共 {rows.Count} 行，其中匹配 {match} 行。");
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (grid.Rows.Count == 0)
                {
                    Log("没有可导出的结果。");
                    return;
                }

                using (var sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                    sfd.FileName = $"对比结果_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    btnExport.Enabled = false;
                    var oldCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        ExportToXlsx(sfd.FileName);
                    }
                    finally
                    {
                        Cursor.Current = oldCursor;
                        btnExport.Enabled = true;
                    }

                    if (MessageBox.Show("导出成功！是否打开所在文件夹？", "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + sfd.FileName + "\""); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 将 grid 的当前内容导出为 .xlsx（两张工作表：结果、参数与统计）
        /// 依赖 NPOI（XSSFWorkbook）
        /// </summary>
        private void ExportToXlsx(string filePath)
        {
            using (var wb = new XSSFWorkbook())
            {
                // ===== 样式 =====
                var boldFont = wb.CreateFont();
                boldFont.IsBold = true;

                ICellStyle headerStyle = wb.CreateCellStyle();
                headerStyle.SetFont(boldFont);
                headerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

                ICellStyle intStyle = wb.CreateCellStyle();
                intStyle.DataFormat = wb.CreateDataFormat().GetFormat("0"); // 整数

                // ===== 工作表1：结果 =====
                var sh = wb.CreateSheet("结果");

                // 表头
                var header = sh.CreateRow(0);
                string[] headers = { "文件夹", "应有", "实际", "差值", "匹配" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var c = header.CreateCell(i);
                    c.SetCellValue(headers[i]);
                    c.CellStyle = headerStyle;
                }
                sh.CreateFreezePane(0, 1);

                // 数据
                int rowIndex = 1;
                int sumExpected = 0, sumActual = 0, sumDiff = 0, matchCount = 0, totalRows = 0;
                foreach (DataGridViewRow dgvr in grid.Rows)
                {
                    if (dgvr.IsNewRow) continue;
                    totalRows++;

                    string folder = dgvr.Cells[0]?.Value?.ToString() ?? "";
                    int expected = SafeToInt(dgvr.Cells[1]?.Value);
                    int actual = SafeToInt(dgvr.Cells[2]?.Value);
                    int diff = SafeToInt(dgvr.Cells[3]?.Value);
                    string match = dgvr.Cells[4]?.Value?.ToString() ?? "";  // "✔" / "✖" 或 True/False

                    sumExpected += expected;
                    sumActual += actual;
                    sumDiff += diff;
                    if (match == "✔" || match.Equals("true", StringComparison.OrdinalIgnoreCase)) matchCount++;

                    var r = sh.CreateRow(rowIndex++);
                    r.CreateCell(0).SetCellValue(folder);

                    var c1 = r.CreateCell(1); c1.SetCellValue(expected); c1.CellStyle = intStyle;
                    var c2 = r.CreateCell(2); c2.SetCellValue(actual); c2.CellStyle = intStyle;
                    var c3 = r.CreateCell(3); c3.SetCellValue(diff); c3.CellStyle = intStyle;

                    r.CreateCell(4).SetCellValue(match);
                }

                // 自动列宽
                for (int i = 0; i < headers.Length; i++)
                    sh.AutoSizeColumn(i);

                // ===== 工作表2：参数与统计 =====
                var sh2 = wb.CreateSheet("参数与统计");
                int r2 = 0;

                void KV(string key, string val)
                {
                    var row = sh2.CreateRow(r2++);
                    var k = row.CreateCell(0); k.SetCellValue(key); k.CellStyle = headerStyle;
                    row.CreateCell(1).SetCellValue(val ?? "");
                }

                // 参数
                KV("导出时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                KV("Excel 文件路径", txtExcel?.Text);
                KV("根目录", txtRoot?.Text);
                // 记录自动识别的列名
                KV("字段映射-文件夹列", _folderColName ?? "");
                KV("字段映射-应有数量列", _expectedColName ?? "");

                // 勾选格式
                var formats = new List<string>();
                foreach (var item in chkFormats.CheckedItems) formats.Add(item.ToString());
                KV("勾选的格式", string.Join(", ", formats));

                // 空一行
                r2++;

                // 统计
                // 这里 sum* 与 matchCount/totalRows 需要在“结果”写入时积累；我们在上面已经累计完毕
                // 为了复用，重新计算一遍会更直观（也可把变量提升到外层，但这里简单处理）
                int sumExpected2 = 0, sumActual2 = 0, sumDiff2 = 0, matchCount2 = 0, totalRows2 = 0;
                foreach (DataGridViewRow dgvr in grid.Rows)
                {
                    if (dgvr.IsNewRow) continue;
                    totalRows2++;
                    int expected = SafeToInt(dgvr.Cells[1]?.Value);
                    int actual = SafeToInt(dgvr.Cells[2]?.Value);
                    int diff = SafeToInt(dgvr.Cells[3]?.Value);
                    string match = dgvr.Cells[4]?.Value?.ToString() ?? "";
                    sumExpected2 += expected;
                    sumActual2 += actual;
                    sumDiff2 += diff;
                    if (match == "✔" || match.Equals("true", StringComparison.OrdinalIgnoreCase)) matchCount2++;
                }

                KV("总行数", totalRows2.ToString());
                KV("匹配行数", matchCount2.ToString());
                KV("应有合计", sumExpected2.ToString());
                KV("实际合计", sumActual2.ToString());
                KV("差值合计", sumDiff2.ToString());

                sh2.AutoSizeColumn(0);
                sh2.AutoSizeColumn(1);

                // ===== 保存 =====
                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fs);
                }
            }
        }

        private static int SafeToInt(object v)
        {
            if (v == null) return 0;
            if (v is int ii) return ii;
            if (int.TryParse(v.ToString(), out int t)) return t;
            return 0;
        }

        private List<CompareTarget> CollectCompareTargets(string root)
        {
            try
            {
                var targets = new List<CompareTarget>();
                foreach (var firstLevelDirectory in Directory.GetDirectories(root))
                {
                    targets.Add(CreateCompareTarget(root, firstLevelDirectory));
                    foreach (var secondLevelDirectory in Directory.GetDirectories(firstLevelDirectory))
                        targets.Add(CreateCompareTarget(root, secondLevelDirectory));
                }
                return targets;
            }
            catch (Exception ex)
            {
                Log($"读取根目录前两层文件夹失败：{ex.Message}");
                return null;
            }
        }

        private static CompareTarget CreateCompareTarget(string root, string directory)
        {
            var rootPrefix = GetRootPathPrefix(root);
            var fullPath = Path.GetFullPath(directory);
            var relativePath = fullPath.Substring(rootPrefix.Length);

            return new CompareTarget
            {
                FullPath = fullPath,
                RelativePath = relativePath,
                Name = Path.GetFileName(fullPath)
            };
        }

        private CompareTargetResolution ResolveCompareTarget(string root, string excelValue, List<CompareTarget> targets)
        {
            if (excelValue.IndexOf('\\') >= 0 || excelValue.IndexOf('/') >= 0)
            {
                string relativePath;
                if (!TryNormalizeRelativePath(root, excelValue, out relativePath))
                {
                    Log($"无效相对路径：{excelValue}");
                    return new CompareTargetResolution { Status = "无效路径" };
                }

                var exactTarget = targets.FirstOrDefault(target =>
                    string.Equals(target.RelativePath, relativePath, StringComparison.OrdinalIgnoreCase));
                if (exactTarget == null)
                {
                    Log($"未找到：{relativePath}");
                    return new CompareTargetResolution { Status = "未找到文件夹" };
                }
                return new CompareTargetResolution { Target = exactTarget };
            }

            var candidates = targets
                .Where(target => string.Equals(target.Name, excelValue, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (candidates.Count == 1) return new CompareTargetResolution { Target = candidates[0] };

            if (candidates.Count == 0)
            {
                Log($"未找到：{excelValue}");
                return new CompareTargetResolution { Status = "未找到文件夹" };
            }
            else
            {
                Log($"名称“{excelValue}”存在多个匹配目录：{string.Join("、", candidates.Select(target => target.RelativePath))}。请在Excel A列填写相对路径进行精确匹配。");
                return new CompareTargetResolution { Status = "名称重复" };
            }
        }

        private static bool TryNormalizeRelativePath(string root, string value, out string relativePath)
        {
            relativePath = null;
            var input = (value ?? string.Empty).Trim();
            if (input.Length == 0 || Path.IsPathRooted(input)) return false;

            var parts = input.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 2 || parts.Any(part => part == "." || part == "..")) return false;

            try
            {
                var rootPrefix = GetRootPathPrefix(root);
                var combinedPath = Path.Combine(rootPrefix, string.Join(Path.DirectorySeparatorChar.ToString(), parts));
                var fullPath = Path.GetFullPath(combinedPath);
                if (!fullPath.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase)) return false;

                relativePath = fullPath.Substring(rootPrefix.Length);
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

        private static int CountActual(string dir, bool useJPG, bool usePNG, bool useTIF, bool useGIF, bool usePDF)
        {
            int actual = 0;
            if (useJPG) actual += CountFiles(dir, new[] { ".jpg", ".jpeg" });
            if (usePNG) actual += CountFiles(dir, new[] { ".png" });
            if (useTIF) actual += CountFiles(dir, new[] { ".tif", ".tiff" });
            if (useGIF) actual += CountFiles(dir, new[] { ".gif" });
            if (usePDF) actual += CountPdfPages(dir); // 注意：PDF 统计页数
            return actual;
        }

        private static int CountFiles(string dir, IEnumerable<string> exts)
        {
            int n = 0;
            foreach (var f in Directory.GetFiles(dir))
            {
                var ext = Path.GetExtension(f);
                foreach (var e in exts)
                {
                    if (ext.Equals(e, StringComparison.OrdinalIgnoreCase)) { n++; break; }
                }
            }
            return n;
        }

        private static int CountPdfPages(string dir)
        {
            int pages = 0;
            foreach (var f in Directory.GetFiles(dir))
            {
                if (!Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    using (var reader = new PdfReader(f))
                    {
                        pages += reader.NumberOfPages;
                    }
                }
                catch { /* 忽略坏文件 */ }
            }
            return pages;
        }
    }
}

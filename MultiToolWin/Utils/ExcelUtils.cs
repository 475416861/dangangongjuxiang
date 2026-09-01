using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // .xlsx
using NPOI.HSSF.UserModel; // .xls
// 文件头需要：
using System.Data;
using System.Data.OleDb;


namespace MultiToolWin.Utils
{
    public static class ExcelUtils
    {
        // ---------- 基础 ----------
        private static IWorkbook OpenWorkbook(string path)
        {
            using (var fs = File.OpenRead(path))
            {
                if (Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    return new HSSFWorkbook(fs);
                return new XSSFWorkbook(fs);
            }
        }

        public static List<string> GetHeader(string path, int sheetIndex = 0)
        {
            var headers = new List<string>();
            using (var fs = File.OpenRead(path))
            using (IWorkbook wb = Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)
                ? (IWorkbook)new HSSFWorkbook(fs) : new XSSFWorkbook(fs))
            {
                var sheet = wb.GetSheetAt(sheetIndex);
                var row0 = sheet.GetRow(sheet.FirstRowNum);
                if (row0 == null) return headers;
                for (int i = row0.FirstCellNum; i < row0.LastCellNum; i++)
                {
                    var cell = row0.GetCell(i);
                    headers.Add(cell?.ToString()?.Trim() ?? $"Col{i + 1}");
                }
            }
            return headers;
        }

        public static List<string> ReadColumn(string path, string columnName, int sheetIndex = 0)
        {
            var result = new List<string>();
            using (var fs = File.OpenRead(path))
            using (IWorkbook wb = Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)
                ? (IWorkbook)new HSSFWorkbook(fs) : new XSSFWorkbook(fs))
            {
                var sheet = wb.GetSheetAt(sheetIndex);
                if (sheet == null) return result;

                var headerRow = sheet.GetRow(sheet.FirstRowNum);
                int colIndex = -1;
                for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
                {
                    var name = headerRow.GetCell(i)?.ToString()?.Trim();
                    if (string.Equals(name, columnName, StringComparison.OrdinalIgnoreCase))
                    { colIndex = i; break; }
                }
                if (colIndex < 0) return result;

                for (int r = sheet.FirstRowNum + 1; r <= sheet.LastRowNum; r++)
                {
                    var row = sheet.GetRow(r);
                    var v = row?.GetCell(colIndex)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) result.Add(v);
                }
            }
            return result;
        }

        // ---------- 批量重命名：导出/读取映射 ----------
        public static void ExportSubfolders(string root, string xlsxPath)
        {
            using (var wb = new XSSFWorkbook())
            {
                var sh = wb.CreateSheet("Folders");
                var header = sh.CreateRow(0);
                header.CreateCell(0).SetCellValue("OldFolder");
                header.CreateCell(1).SetCellValue("NewFolder");

                // 取顶层子文件夹并按“自然排序”排名字：1,2,3,10,200...
                var dirs = Directory.GetDirectories(root);
                Array.Sort(dirs, (a, b) =>
                    new NaturalStringComparer().Compare(Path.GetFileName(a), Path.GetFileName(b)));

                int r = 1;
                foreach (var d in dirs)
                {
                    var row = sh.CreateRow(r++);
                    row.CreateCell(0).SetCellValue(Path.GetFileName(d));
                }

                using (var fs = File.Create(xlsxPath))
                {
                    wb.Write(fs);
                }
            }
        }

        // 放在 ExcelUtils 类里（或同命名空间下），自然排序比较器
        internal class NaturalStringComparer : IComparer<string>
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


        public static List<(string OldFolder, string NewFolder)> ReadFolderMap(string xlsxPath)
        {
            var list = new List<(string, string)>();
            using (var wb = OpenWorkbook(xlsxPath))
            {
                var sh = wb.GetSheetAt(0);
                var header = sh.GetRow(sh.FirstRowNum);
                int colOld = -1, colNew = -1;
                for (int i = header.FirstCellNum; i < header.LastCellNum; i++)
                {
                    var name = header.GetCell(i)?.ToString()?.Trim();
                    if (string.Equals(name, "OldFolder", StringComparison.OrdinalIgnoreCase)) colOld = i;
                    if (string.Equals(name, "NewFolder", StringComparison.OrdinalIgnoreCase)) colNew = i;
                }
                if (colOld < 0 || colNew < 0) return list;

                for (int r = sh.FirstRowNum + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    var o = row?.GetCell(colOld)?.ToString()?.Trim();
                    var n = row?.GetCell(colNew)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(o) && !string.IsNullOrEmpty(n))
                        list.Add((o, n));
                }
            }
            return list;
        }

        // ---------- 数量对比：读取两列 ----------
        public static List<(string Folder, int Expected)> ReadExpected(string xlsxPath, string folderCol, string expectedCol)
        {
            var res = new List<(string, int)>();
            using (var wb = OpenWorkbook(xlsxPath))
            {
                var sh = wb.GetSheetAt(0);
                var header = sh.GetRow(sh.FirstRowNum);
                int colF = -1, colE = -1;
                for (int i = header.FirstCellNum; i < header.LastCellNum; i++)
                {
                    var name = header.GetCell(i)?.ToString()?.Trim();
                    if (string.Equals(name, folderCol, StringComparison.OrdinalIgnoreCase)) colF = i;
                    if (string.Equals(name, expectedCol, StringComparison.OrdinalIgnoreCase)) colE = i;
                }
                if (colF < 0 || colE < 0) return res;

                for (int r = sh.FirstRowNum + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    var f = row?.GetCell(colF)?.ToString()?.Trim();
                    var eStr = row?.GetCell(colE)?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(f)) continue;
                    int expected = 0;
                    int.TryParse(eStr, out expected);
                    res.Add((f, expected));
                }
            }
            return res;
        }

        // ---------- Excel 图片提取：读取“提取计划” ----------
        public class ExtractPlanRow
        {
            public string Src;   // 原文件夹名字
            public string Dest;  // 新文件夹名字
            public int Count;    // 页数/数量（<=0 表示“全取剩余”）
        }

        public static List<ExtractPlanRow> ReadExtractPlan(string xlsxPath)
        {
            var plan = new List<ExtractPlanRow>();
            using (var fs = File.OpenRead(xlsxPath))
            using (IWorkbook wb = Path.GetExtension(xlsxPath).Equals(".xls", StringComparison.OrdinalIgnoreCase)
                ? (IWorkbook)new HSSFWorkbook(fs) : new XSSFWorkbook(fs))
            {
                var sh = wb.GetSheetAt(0);
                if (sh == null) return plan;
                var header = sh.GetRow(sh.FirstRowNum);
                if (header == null) return plan;

                string[] srcCands = { "原文件夹名字", "原文件夹名称", "源文件夹", "OldFolder", "Folder", "Name" };
                string[] destCands = { "新文件夹名字", "新文件夹名称", "目标文件夹", "NewFolder", "Dest", "Target" };
                string[] cntCands = { "页数", "数量", "张数", "Count", "Num" };

                int colSrc = -1, colDst = -1, colCnt = -1;
                for (int i = header.FirstCellNum; i < header.LastCellNum; i++)
                {
                    var h = header.GetCell(i)?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(h)) continue;
                    if (colSrc < 0 && Array.Exists(srcCands, s => s.Equals(h, StringComparison.OrdinalIgnoreCase))) colSrc = i;
                    if (colDst < 0 && Array.Exists(destCands, s => s.Equals(h, StringComparison.OrdinalIgnoreCase))) colDst = i;
                    if (colCnt < 0 && Array.Exists(cntCands, s => s.Equals(h, StringComparison.OrdinalIgnoreCase))) colCnt = i;
                }
                if (colSrc < 0 || colDst < 0) return plan;

                for (int r = sh.FirstRowNum + 1; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    if (row == null) continue;
                    var src = row.GetCell(colSrc)?.ToString()?.Trim();
                    var dest = row.GetCell(colDst)?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(dest)) continue;

                    int cnt = 0;
                    if (colCnt >= 0) int.TryParse(row.GetCell(colCnt)?.ToString()?.Trim(), out cnt);
                    plan.Add(new ExtractPlanRow { Src = src, Dest = dest, Count = cnt });
                }
            }
            return plan;
        }
        // 读取 A 列（第一列）；支持 .xlsx/.xls（用 NPOI），兼容 .csv/.txt（逐行）
        // 不做任何“名称清洗”，只 Trim；空行会被忽略。
        public static List<string> ReadFolderNames(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("文件不存在", path);

            string ext = Path.GetExtension(path).ToLowerInvariant();

            // 文本/CSV：每行一个名称
            if (ext == ".csv" || ext == ".txt")
                return new List<string>(File.ReadAllLines(path));

            if (ext != ".xlsx" && ext != ".xls")
                throw new NotSupportedException("仅支持 .xlsx / .xls / .csv / .txt");

            List<string> list = new List<string>();

            using (var wb = OpenWorkbook(path))               // 你已有的 NPOI 打开函数
            {
                var sh = wb.GetSheetAt(0);
                if (sh == null) return list;

                // 判断第一行是否是表头（包含“名/名称/folder/name”等），是则跳过第一行
                int startRow = sh.FirstRowNum;
                try
                {
                    var r0 = sh.GetRow(sh.FirstRowNum);
                    string h = r0?.GetCell(0)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(h))
                    {
                        string u = h.ToUpperInvariant();
                        if (u.Contains("名") || u.Contains("NAME") || u.Contains("FOLDER"))
                            startRow = sh.FirstRowNum + 1;
                    }
                }
                catch { /* 安全兜底 */ }

                for (int r = startRow; r <= sh.LastRowNum; r++)
                {
                    var row = sh.GetRow(r);
                    if (row == null) continue;
                    var cell = row.GetCell(0);                // A 列
                    string v = cell?.ToString();
                    if (!string.IsNullOrWhiteSpace(v))
                        list.Add(v.Trim());
                }
            }

            return list;
        }

        // 兼容名（PageMkFolders 也会查找这个）
        public static List<string> ReadFirstColumn(string path)
        {
            return ReadFolderNames(path);
        }

    }
}


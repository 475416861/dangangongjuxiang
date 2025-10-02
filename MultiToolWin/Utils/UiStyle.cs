using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace MultiToolWin.Utils
{
    public static class UiStyle
    {
        // 全局字体（Win7有）
        public static readonly Font DefaultFont = new Font("Microsoft YaHei UI", 10f);

        // 顶部“文本框 + 浏览按钮”
        public static void StyleBrowseRow(TextBox tb, Button btn)
        {
            if (tb != null)
            {
                tb.Margin = new Padding(0, 0, 6, 0);
                tb.Anchor = AnchorStyles.Left | AnchorStyles.Right;
                tb.Font = DefaultFont;
            }
            if (btn != null)
            {
                btn.Width = 80;
                btn.Height = 28;
                btn.Margin = new Padding(6, 0, 0, 0);
                btn.FlatStyle = FlatStyle.System;     // Win7 稳定
                btn.Font = DefaultFont;
            }
        }

        // “字段映射”标签/下拉
        public static void StyleMapLabel(Label lbl, int topPad = 6)
        {
            if (lbl == null) return;
            lbl.Margin = new Padding(0, topPad, 6, 0);
            lbl.Font = DefaultFont;
        }
        public static void StyleMapCombo(ComboBox cb, int rightPad = 12)
        {
            if (cb == null) return;
            cb.DropDownStyle = ComboBoxStyle.DropDownList;
            cb.Width = cb.Width < 150 ? 150 : cb.Width;
            cb.Margin = new Padding(0, 3, rightPad, 0);
            cb.Font = DefaultFont;
        }

        // 复选列表
        public static void StyleCheckedList(CheckedListBox clb, int height = 94)
        {
            if (clb == null) return;
            clb.CheckOnClick = true;
            clb.IntegralHeight = false;
            clb.BorderStyle = BorderStyle.FixedSingle;
            clb.ItemHeight = 22;
            clb.Height = height;
            clb.Font = DefaultFont;
        }

        // DataGridView（Win7 兼容）
        public static void StyleGrid(DataGridView grid)
        {
            if (grid == null) return;

            // 开启双缓冲，减少闪烁（反射）
            try
            {
                typeof(DataGridView)
                    .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.SetValue(grid, true, null);
            }
            catch { }

            grid.Font = DefaultFont;
            grid.BackgroundColor = Color.White;
            grid.GridColor = Color.FromArgb(230, 235, 240);

            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(243, 246, 249);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 35, 40);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersHeight = 30;

            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(227, 239, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;

            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);

            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.RowHeadersVisible = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        public static void AlignRight(params DataGridViewColumn[] cols)
        {
            if (cols == null) return;
            foreach (var c in cols) if (c != null)
                    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }
        public static void AlignCenter(params DataGridViewColumn[] cols)
        {
            if (cols == null) return;
            foreach (var c in cols) if (c != null)
                    c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
    }
}

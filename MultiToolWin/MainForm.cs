using System;
using System.Drawing;
using System.Windows.Forms;
using MultiToolWin.Pages;

namespace MultiToolWin
{
    public class MainForm : Form
    {
        private ListBox navList;
        private Panel contentPanel;
        private TextBox txtLog;
        private Button btnClearLog, btnCopyLog;

        // 页面
        private PageExtract pageExtract;
        private PageRename pageRename;
        private PageCompare pageCompare;
        private PageMkFolders pageMkFolders;
     

        // —— 导航自绘需要的字段（类级别） ——
        private int _hoverIndex = -1;

        private static readonly Color ColSelectedBack = ColorTranslator.FromHtml("#EDEFF3"); // 选中底
        private static readonly Color ColHoverBack = ColorTranslator.FromHtml("#F2F4F7");   // 悬停底
        private static readonly Color ColText = ColorTranslator.FromHtml("#1F2328");        // 文本色
        private static readonly Color ColIndicator = ColorTranslator.FromHtml("#3B82F6");   // 左2px指示条

        public MainForm()
        {
            Text = "整合工具箱 v3.1  (.NET 4.7.2)";
            Width = 1000;
            Height = 650;
            StartPosition = FormStartPosition.CenterScreen;

            // 左侧导航
            navList = new ListBox
            {
                Dock = DockStyle.Left,
                Width = 120
            };
            navList.Items.AddRange(new object[] {
                "Excel 图片提取",
                "批量重命名",
                "数量/页数对比",
                "批量建文件夹",
            
            });
            navList.SelectedIndexChanged += (s, e) => SwitchPage(navList.SelectedIndex);

            // —— 极简外观设置 ——
            navList.Font = new Font("Microsoft YaHei UI", 10f);
            navList.BorderStyle = BorderStyle.None;
            navList.IntegralHeight = false;
            navList.ItemHeight = 30;                 // 28~32 均可
            navList.DrawMode = DrawMode.OwnerDrawFixed;
            navList.BackColor = Color.White;

            // 事件：自绘 + 悬停
            navList.DrawItem += NavList_DrawItem;
            navList.MouseMove += (s, e) =>
            {
                int idx = navList.IndexFromPoint(e.Location);
                if (idx != _hoverIndex) { _hoverIndex = idx; navList.Invalidate(); }
            };
            navList.MouseLeave += (s, e) => { _hoverIndex = -1; navList.Invalidate(); };

            // 右侧内容区
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12)
            };

            // 底部全局日志
            var logPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 150
            };
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
            var logButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                FlowDirection = FlowDirection.RightToLeft
            };
            btnClearLog = new Button { Text = "清空", Width = 80, Height = 28 };
            btnCopyLog = new Button { Text = "复制", Width = 80, Height = 28 };
            btnClearLog.Click += (s, e) => txtLog.Clear();
            btnCopyLog.Click += (s, e) => { if (!string.IsNullOrEmpty(txtLog.Text)) Clipboard.SetText(txtLog.Text); };
            logButtons.Controls.Add(btnClearLog);
            logButtons.Controls.Add(btnCopyLog);
            logPanel.Controls.Add(txtLog);
            logPanel.Controls.Add(logButtons);

            Controls.Add(contentPanel);
            Controls.Add(logPanel);
            Controls.Add(navList);

            // 创建页面
            Action<string> logger = Log;
            pageExtract = new PageExtract(logger);
            pageRename = new PageRename(logger);
            pageCompare = new PageCompare(logger);
            pageMkFolders = new PageMkFolders(logger);
         

            // 默认页
            navList.SelectedIndex = 0;
        }

        private void SwitchPage(int index)
        {
            contentPanel.Controls.Clear();
            UserControl page = null;
            switch (index)
            {
                case 0: page = pageExtract; break;
                case 1: page = pageRename; break;
                case 2: page = pageCompare; break;
                case 3: page = pageMkFolders; break;
              
            }
            if (page != null)
            {
                page.Dock = DockStyle.Fill;
                contentPanel.Controls.Add(page);
            }
        }

        public void Log(string msg)
        {
            if (txtLog == null || txtLog.IsDisposed) return;

            if (txtLog.InvokeRequired)
            {
                try { txtLog.BeginInvoke(new Action<string>(Log), msg); }
                catch { }
                return;
            }

            var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
            try { txtLog.AppendText(line + Environment.NewLine); } catch { }
        }

        // —— ListBox 极简自绘 ——
        private void NavList_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0 || e.Index >= navList.Items.Count) return;

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool hovered = (!selected && e.Index == _hoverIndex);

            // 背景
            using (var back = new SolidBrush(selected ? ColSelectedBack : hovered ? ColHoverBack : Color.White))
                e.Graphics.FillRectangle(back, e.Bounds);

            // 左侧 2px 指示条（仅选中）
            if (selected)
            {
                var bar = new Rectangle(e.Bounds.X, e.Bounds.Y, 2, e.Bounds.Height);
                using (var sb = new SolidBrush(ColIndicator))
                    e.Graphics.FillRectangle(sb, bar);
            }

            // 文本区域：左右留白
            var textRect = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 20, e.Bounds.Height);

            TextRenderer.DrawText(
                e.Graphics,
                navList.Items[e.Index].ToString(),
                navList.Font,
                textRect,
                ColText,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }
    }
}

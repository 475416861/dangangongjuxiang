using System;
using System.Drawing;
using System.Windows.Forms;
using MultiToolWin.Pages;

namespace MultiToolWin
{
    public class MainForm : Form
    {
        private Button[] navButtons;
        private Font navRegularFont;
        private Font navCurrentFont;
        private Panel contentPanel;
        private TextBox txtLog;
        private Button btnClearLog, btnCopyLog;

        // 页面
        private PageExtract pageExtract;
        private PageRename pageRename;
        private PageCompare pageCompare;
        private PageMkFolders pageMkFolders;
        public MainForm()
        {
            Text = "整合工具箱 v3.2  (.NET 4.7.2)";
            Width = 1000;
            Height = 650;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9f);
            navRegularFont = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular);
            navCurrentFont = new Font("Microsoft YaHei UI", 9f, FontStyle.Bold);

            // 左侧导航
            var navigationPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 142,
                Padding = new Padding(8, 10, 8, 8)
            };
            var navigationLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 4
            };
            string[] pageNames = { "Excel 图片提取", "批量重命名", "数量/页数对比", "批量建文件夹" };
            navButtons = new Button[pageNames.Length];
            for (int i = 0; i < pageNames.Length; i++)
            {
                int pageIndex = i;
                var button = new Button
                {
                    Text = pageNames[i],
                    Dock = DockStyle.Fill,
                    Height = 32,
                    Margin = new Padding(0, 0, 0, 6),
                    FlatStyle = FlatStyle.System,
                    Font = navRegularFont
                };
                button.Click += (s, e) => SwitchPage(pageIndex);
                navButtons[i] = button;
                navigationLayout.Controls.Add(button, 0, i);
                navigationLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }
            navigationPanel.Controls.Add(navigationLayout);

            // 右侧内容区
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12)
            };

            // 底部全局日志
            var logPanel = new GroupBox
            {
                Dock = DockStyle.Bottom,
                Height = 158,
                Text = "运行日志",
                Padding = new Padding(8)
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
                Height = 34,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 0, 0, 2)
            };
            btnClearLog = new Button { Text = "清空", Width = 80, Height = 26, FlatStyle = FlatStyle.System };
            btnCopyLog = new Button { Text = "复制", Width = 80, Height = 26, FlatStyle = FlatStyle.System };
            btnClearLog.Click += (s, e) => txtLog.Clear();
            btnCopyLog.Click += (s, e) => { if (!string.IsNullOrEmpty(txtLog.Text)) Clipboard.SetText(txtLog.Text); };
            logButtons.Controls.Add(btnClearLog);
            logButtons.Controls.Add(btnCopyLog);
            logPanel.Controls.Add(txtLog);
            logPanel.Controls.Add(logButtons);

            Controls.Add(contentPanel);
            Controls.Add(logPanel);
            Controls.Add(navigationPanel);

            // 创建页面
            Action<string> logger = Log;
            pageExtract = new PageExtract(logger);
            pageRename = new PageRename(logger);
            pageCompare = new PageCompare(logger);
            pageMkFolders = new PageMkFolders(logger);
         

            // 默认页
            SwitchPage(0);
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

            for (int i = 0; i < navButtons.Length; i++)
            {
                navButtons[i].Enabled = true;
                navButtons[i].Font = i == index ? navCurrentFont : navRegularFont;
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

    }
}

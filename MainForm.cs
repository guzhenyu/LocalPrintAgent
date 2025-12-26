using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace LocalPrintAgent
{
    public partial class MainForm : Form
    {
        private readonly NotifyIcon _tray;
        private readonly PrintServer _server;
        private readonly AppConfig _config;

        public MainForm()
        {
            // InitializeComponent();

            // 窗体隐藏：不显示任务栏、不闪一下
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Opacity = 0;

            _config = AppConfig.LoadOrCreate();

            _server = new PrintServer(
                prefix: "http://127.0.0.1:9123/",
                config: _config,
                logger: Log
            );

            _tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "LocalPrintAgent (静默打印)"
            };

            var menu = new ContextMenuStrip();
            menu.Items.Add("查看 Token", null, (_, __) =>
            {
                Clipboard.SetText(_config.Token);
                MessageBox.Show(
                    $"Token 已复制到剪贴板：\n\n{_config.Token}\n\n前端请求需带 X-Print-Token 头。",
                    "LocalPrintAgent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            });

            menu.Items.Add("打印机配置", null, (_, __) =>
            {
                using var dialog = new PrinterConfigForm(_config);
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _config.A3PrinterName = dialog.A3PrinterName;
                    _config.A4PrinterName = dialog.A4PrinterName;
                    _config.Save();
                }
            });

            menu.Items.Add("打开日志目录", null, (_, __) =>
            {
                var dir = AppConfig.GetAppDir();
                Directory.CreateDirectory(dir);
                System.Diagnostics.Process.Start("explorer.exe", dir);
            });

            menu.Items.Add("退出", null, (_, __) => Close());
            _tray.ContextMenuStrip = menu;

            // 启动 HTTP 服务
            try
            {
                _server.Start();
                _tray.BalloonTipTitle = "LocalPrintAgent 已启动";
                _tray.BalloonTipText = "监听 127.0.0.1:9123，等待打印请求";
                _tray.ShowBalloonTip(1500);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "启动失败：\n" + ex.Message + "\n\n常见原因：端口被占用，或 HttpListener URL 预留问题。",
                    "LocalPrintAgent",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                throw;
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _server.Stop();
            _tray.Visible = false;
            _tray.Dispose();
            base.OnFormClosing(e);
        }

        private void Log(string line)
        {
            try
            {
                AppConfig.AppendLog(line);
            }
            catch { /* ignore */ }
        }
    }
}

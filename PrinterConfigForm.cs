using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace LocalPrintAgent
{
    internal sealed class PrinterConfigForm : Form
    {
        private const string Placeholder = "(请选择)";
        private readonly ComboBox _a3Combo;
        private readonly ComboBox _a4Combo;
        private readonly Button _okButton;

        public string A3PrinterName => GetSelectedPrinter(_a3Combo);
        public string A4PrinterName => GetSelectedPrinter(_a4Combo);

        public PrinterConfigForm(AppConfig config)
        {
            Text = "打印机配置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;

            _a3Combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };
            _a4Combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Dock = DockStyle.Fill
            };

            var printers = new List<string>();
            foreach (string p in PrinterSettings.InstalledPrinters) printers.Add(p);

            FillCombo(_a3Combo, printers, config.A3PrinterName);
            FillCombo(_a4Combo, printers, config.A4PrinterName);

            var labelA3 = new Label
            {
                Text = "A3 打印机：",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Left
            };
            var labelA4 = new Label
            {
                Text = "A4 打印机：",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Left
            };

            _okButton = new Button { Text = "保存", DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(_okButton);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            layout.Controls.Add(labelA3, 0, 0);
            layout.Controls.Add(_a3Combo, 1, 0);
            layout.Controls.Add(labelA4, 0, 1);
            layout.Controls.Add(_a4Combo, 1, 1);
            layout.SetColumnSpan(buttons, 2);
            layout.Controls.Add(buttons, 0, 2);

            Controls.Add(layout);
            ClientSize = new Size(ClientSize.Width * 2, ClientSize.Height);
            AcceptButton = _okButton;
            CancelButton = cancelButton;

            if (printers.Count == 0)
            {
                _okButton.Enabled = false;
                _a3Combo.Enabled = false;
                _a4Combo.Enabled = false;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                if (!IsValidSelection(_a3Combo) || !IsValidSelection(_a4Combo))
                {
                    MessageBox.Show(
                        "请为 A3 和 A4 选择可用的打印机。",
                        "打印机配置",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    e.Cancel = true;
                }
            }

            base.OnFormClosing(e);
        }

        private static void FillCombo(ComboBox combo, List<string> printers, string selected)
        {
            combo.Items.Add(Placeholder);
            foreach (var p in printers) combo.Items.Add(p);

            if (!string.IsNullOrWhiteSpace(selected) && combo.Items.Contains(selected))
            {
                combo.SelectedItem = selected;
            }
            else
            {
                combo.SelectedIndex = 0;
            }
        }

        private static bool IsValidSelection(ComboBox combo)
        {
            var value = combo.SelectedItem as string;
            return !string.IsNullOrWhiteSpace(value)
                && !string.Equals(value, Placeholder, StringComparison.Ordinal);
        }

        private static string GetSelectedPrinter(ComboBox combo)
        {
            var value = combo.SelectedItem as string;
            return IsValidSelection(combo) ? value ?? "" : "";
        }
    }
}

using System;
using System.IO;
using System.Windows.Forms;

namespace VectraConnect.UI
{
    public class SettingsDialog : Form
    {
        private TextBox  _folderBox;
        private CheckBox _csvCheck;
        private Button   _browseBtn;
        private Button   _okBtn;
        private Button   _cancelBtn;

        public SettingsDialog()
        {
            Text            = "Vectra Connect — Settings";
            Width           = 520;
            Height          = 200;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition   = FormStartPosition.CenterScreen;
            MaximizeBox     = false;
            MinimizeBox     = false;

            // Output folder row
            var folderLabel = new Label { Text = "Output folder:", Left = 12, Top = 20, Width = 100, AutoSize = true };

            _folderBox = new TextBox
            {
                Left  = 120, Top  = 17,
                Width = 300, Text = SettingsManager.OutputFolder
            };

            _browseBtn = new Button
            {
                Text  = "Browse…",
                Left  = 428, Top  = 15,
                Width = 68,  Height = 24
            };
            _browseBtn.Click += BrowseBtn_Click;

            // CSV checkbox
            _csvCheck = new CheckBox
            {
                Text    = "Also export CSV files (one per sheet)",
                Left    = 120, Top = 55,
                Width   = 280,
                Checked = SettingsManager.IncludeCsv
            };

            // OK / Cancel
            _okBtn = new Button
            {
                Text         = "OK",
                DialogResult = DialogResult.OK,
                Left         = 320, Top = 110,
                Width        = 80
            };
            _okBtn.Click += OkBtn_Click;

            _cancelBtn = new Button
            {
                Text         = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left         = 416, Top = 110,
                Width        = 80
            };

            AcceptButton = _okBtn;
            CancelButton = _cancelBtn;

            Controls.AddRange(new Control[] {
                folderLabel, _folderBox, _browseBtn, _csvCheck, _okBtn, _cancelBtn
            });
        }

        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description         = "Select default output folder";
                dlg.ShowNewFolderButton = true;
                if (Directory.Exists(_folderBox.Text))
                    dlg.SelectedPath = _folderBox.Text;

                if (dlg.ShowDialog() == DialogResult.OK)
                    _folderBox.Text = dlg.SelectedPath;
            }
        }

        private void OkBtn_Click(object sender, EventArgs e)
        {
            SettingsManager.OutputFolder = _folderBox.Text.Trim();
            SettingsManager.IncludeCsv   = _csvCheck.Checked;
        }
    }
}

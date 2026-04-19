using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using VectraConnect.UI;

// IRibbonExtensibility requires the COM-visible attribute
[assembly: System.Runtime.InteropServices.ComVisible(true)]

namespace VectraConnect.Ribbon
{
    [ComVisible(true)]
    public class PublisherRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        // ── IRibbonExtensibility ──────────────────────────────────────────

        public string GetCustomUI(string ribbonID)
        {
            // Load the embedded ribbon XML resource
            return GetResourceText("VectraConnect.Ribbon.PublisherRibbon.xml");
        }

        // ── Ribbon lifecycle ──────────────────────────────────────────────

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // ── Button: Publish ───────────────────────────────────────────────

        public void OnPublishClick(IRibbonControl control)
        {
            try
            {
                var app     = Globals.ThisAddIn.Application;
                var project = app.ActiveProject;

                if (project == null)
                {
                    MessageBox.Show("No project is currently open.",
                                    "Vectra Connect",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                // Determine output folder
                string outputFolder = SettingsManager.OutputFolder;

                if (string.IsNullOrWhiteSpace(outputFolder))
                {
                    outputFolder = PickFolder("Choose output folder for schema files");
                    if (outputFolder == null) return; // user cancelled
                    SettingsManager.OutputFolder = outputFolder;
                }

                // Parse + export
                var schema = ProjectParser.Parse(project);
                bool includeCsv = SettingsManager.IncludeCsv;
                var result = SchemaExporter.Export(schema, outputFolder, includeCsv);

                // Success dialog
                string msg = $"✅ Schema published!\n\n" +
                             $"Excel:  {result.XlsxPath}\n";
                if (includeCsv)
                    msg += $"\nCSV files also written to:\n{outputFolder}";

                var dlgResult = MessageBox.Show(
                    msg + "\n\nOpen output folder?",
                    "Vectra Connect",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (dlgResult == DialogResult.Yes)
                    System.Diagnostics.Process.Start("explorer.exe", outputFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n\n{ex.Message}",
                                "Vectra Connect",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        // ── Button: Settings ──────────────────────────────────────────────

        public void OnSettingsClick(IRibbonControl control)
        {
            using (var dlg = new SettingsDialog())
            {
                dlg.ShowDialog();
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────

        private static string PickFolder(string description)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description         = description;
                dlg.ShowNewFolderButton = true;

                string last = SettingsManager.OutputFolder;
                if (!string.IsNullOrWhiteSpace(last) && Directory.Exists(last))
                    dlg.SelectedPath = last;

                return dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
            }
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException($"Ribbon XML resource '{resourceName}' not found.");
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}

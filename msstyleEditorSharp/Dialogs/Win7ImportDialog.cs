using libmsstyle.Import;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace msstyleEditor
{
    /// <summary>Shows the outcome of a win7 import: status, warnings, and the full markdown report.</summary>
    public class Win7ImportDialog : Form
    {
        public Win7ImportDialog(ImportReport report, string reportPath)
        {
            bool ok = report.Status == ImportStatus.ReadyForVmTest;
            Text = ok ? "Windows 7 style imported" : "Windows 7 import blocked";
            Size = new Size(900, 640);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;

            int applied = report.Counts.Where(kv => kv.Key.StartsWith("applied") && kv.Key != "applied-numeric-part" || kv.Key == "added" || kv.Key == "removed-color").Sum(kv => kv.Value);
            int kept = report.Counts.Where(kv => kv.Key.StartsWith("kept-target")).Sum(kv => kv.Value);
            int missing;
            report.Counts.TryGetValue("skipped-target-missing", out missing);

            var summary = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 70,
                Padding = new Padding(8),
                Text = ok
                    ? $"{applied} properties taken from the Windows 7 style, {kept} kept because the template adapted them for Windows 10, " +
                      $"{missing} have no Windows 10 counterpart. DWM frames: {report.Dwm.Parts.Count(p => p.Method == "name" || p.Method == "alias")} sprites mapped, {report.Dwm.FallbackParts.Count} kept from the template.\r\n" +
                      "The result is now open but not saved. Offline checks passed; that isn't a render test, so try it in a VM before using it."
                    : "Nothing was changed. " + string.Join(" ", report.Errors)
            };

            var log = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new Font(FontFamily.GenericMonospace, 9f),
                Text = ImportReportWriter.ToMarkdown(report).Replace("\r\n", "\n").Replace("\n", Environment.NewLine)
            };

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            var close = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.OK };
            var folder = new Button { Text = "Open report folder", AutoSize = true };
            folder.Click += (s, e) => Process.Start("explorer", "/select,\"" + reportPath + "\"");
            buttons.Controls.Add(close);
            buttons.Controls.Add(folder);
            AcceptButton = close;

            Controls.Add(log);
            // focus Close once shown; setting ActiveControl in the constructor throws because the button isn't visible yet.
            Shown += (s, e) => close.Focus();
            Controls.Add(buttons);
            Controls.Add(summary);
        }
    }
}

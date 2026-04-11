using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace msstyleEditor
{

    public partial class ControlDemoDialog : Form
    {
        private const int PBM_SETSTATE = 0x410;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

        public ControlDemoDialog()
        {
            InitializeComponent();

            // Apply system colors from current msstyles theme
            ApplySystemColors();

            // Hook up Paint events for disabled controls to use proper system colors
            HookDisabledControlsPaintEvents();

            // Native main menu
            MenuItem miClose = new MenuItem("Close",
                (s, e) => this.Close());
            MenuItem miToolWindow = new MenuItem("Toggle Tool Window",
                (s, e) => this.FormBorderStyle = this.FormBorderStyle == FormBorderStyle.Sizable
                    ? FormBorderStyle.SizableToolWindow
                    : FormBorderStyle.Sizable);
            MenuItem miFileDlg = new MenuItem("File Dialog",
                (s, e) => new OpenFileDialog().ShowDialog());
            MenuItem miColorDlg = new MenuItem("Color Dialog",
                (s, e) => new ColorDialog() { FullOpen = true }.ShowDialog());

            var mainMenu = new MainMenu();
            mainMenu.MenuItems.Add(new MenuItem("Window", new MenuItem[] { miClose, miToolWindow }));
            mainMenu.MenuItems.Add(new MenuItem("View", new MenuItem[] { miFileDlg, miColorDlg }));
            this.Menu = mainMenu;

            // Native status bar
            StatusBarPanel panel1 = new StatusBarPanel();
            panel1.Text = "Status Bar";
            StatusBarPanel panel2 = new StatusBarPanel();
            panel2.Text = "Panel 1";
            StatusBar statusBar = new StatusBar();
            statusBar.BackColor = SystemColors.Control;
            statusBar.ShowPanels = true;
            statusBar.Panels.Add(panel1);
            statusBar.Panels.Add(panel2);
            this.Controls.Add(statusBar);

            // List view
            string[] text = new string[] { "Item 1", "Item 2" };
            listView1.Items.Add(new ListViewItem(text, 0, listView1.Groups[0]));
            listView1.Items.Add(new ListViewItem(text, 0, listView1.Groups[1]));

            // Progress bar with states
            SendMessage(progressBarError.Handle, PBM_SETSTATE, 2, 0);
            SendMessage(progressBarPaused.Handle, PBM_SETSTATE, 3, 0);

            // Combo box setup
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }

        private void HookDisabledControlsPaintEvents()
        {
            // Use FlatStyle.System for native msstyle rendering
            // The disabled text color will use the current theme's GRAYTEXT color
            button2.FlatStyle = FlatStyle.System;

            radioButton2.FlatStyle = FlatStyle.System;

            checkBox2.FlatStyle = FlatStyle.System;
            checkBox4.FlatStyle = FlatStyle.System;
            checkBox6.FlatStyle = FlatStyle.System;
        }

        private void DrawClassicCheckBox(PaintEventArgs e, CheckBox cb, Rectangle glyphRect)
        {
            // Fallback to classic rendering
            using (Brush brush = new SolidBrush(SystemColors.Control))
            {
                e.Graphics.FillRectangle(brush, cb.ClientRectangle);
            }

            ButtonState state = ButtonState.Inactive;
            if (cb.Checked)
                state = ButtonState.Checked | ButtonState.Inactive;

            ControlPaint.DrawCheckBox(e.Graphics, glyphRect, state);
        }

        private void ApplySystemColors()
        {
            // Form background
            this.BackColor = SystemColors.Control;

            // Labels - use system colors
            foreach (Control ctrl in this.Controls.OfType<Label>())
            {
                ctrl.BackColor = SystemColors.Control;
                ctrl.ForeColor = SystemColors.ControlText;
            }

            // Buttons - UseVisualStyleBackColor is already set in designer, which handles theming
            // Radio buttons - UseVisualStyleBackColor is already set in designer
            // Checkboxes - UseVisualStyleBackColor is already set in designer
            // Tab pages - UseVisualStyleBackColor is already set in designer

            // TextBox controls
            textBox1.BackColor = SystemColors.Window;
            textBox1.ForeColor = SystemColors.WindowText;
            textBox2.BackColor = SystemColors.Window;
            textBox2.ForeColor = SystemColors.WindowText;
            textBox3.BackColor = SystemColors.Window;
            textBox3.ForeColor = SystemColors.WindowText;
            textBox4.BackColor = SystemColors.Window;
            textBox4.ForeColor = SystemColors.WindowText;

            // ComboBox controls
            comboBox1.BackColor = SystemColors.Window;
            comboBox1.ForeColor = SystemColors.WindowText;
            comboBox2.BackColor = SystemColors.Window;
            comboBox2.ForeColor = SystemColors.WindowText;
            comboBox3.BackColor = SystemColors.Window;
            comboBox3.ForeColor = SystemColors.WindowText;

            // GroupBox
            groupBox1.BackColor = SystemColors.Control;
            groupBox1.ForeColor = SystemColors.ControlText;

            // NumericUpDown
            numericUpDown1.BackColor = SystemColors.Window;
            numericUpDown1.ForeColor = SystemColors.WindowText;

            // LinkLabel
            linkLabel1.LinkColor = SystemColors.HotTrack;
            linkLabel1.VisitedLinkColor = SystemColors.HotTrack;
            linkLabel1.ActiveLinkColor = SystemColors.HotTrack;

            // ScrollBars
            hScrollBar1.BackColor = SystemColors.Control;
            vScrollBar1.BackColor = SystemColors.Control;

            // ProgressBars
            progressBarNormal.BackColor = SystemColors.Control;
            progressBarError.BackColor = SystemColors.Control;
            progressBarPaused.BackColor = SystemColors.Control;
            progressBarMarquee.BackColor = SystemColors.Control;

            // ListBox
            listBox1.BackColor = SystemColors.Window;
            listBox1.ForeColor = SystemColors.WindowText;

            // ListView
            listView1.BackColor = SystemColors.Window;
            listView1.ForeColor = SystemColors.WindowText;

            // TreeView
            treeView1.BackColor = SystemColors.Window;
            treeView1.ForeColor = SystemColors.WindowText;

            // TrackBar
            trackBar1.BackColor = SystemColors.Control;

            // TabControl
            tabControl1.BackColor = SystemColors.Control;
            tabControl1.ForeColor = SystemColors.ControlText;
        }

    }
}

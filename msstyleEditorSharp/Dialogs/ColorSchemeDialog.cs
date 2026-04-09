using libmsstyle;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace msstyleEditor
{
    public partial class ColorSchemeDialog : Form
    {
        public XpColorScheme SelectedScheme { get; private set; }

        public ColorSchemeDialog(List<XpColorScheme> schemes)
        {
            InitializeComponent();

            foreach (var scheme in schemes)
            {
                listBoxSchemes.Items.Add(scheme);
            }

            listBoxSchemes.DisplayMember = "DisplayName";

            if (listBoxSchemes.Items.Count > 0)
                listBoxSchemes.SelectedIndex = 0;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                SelectedScheme = listBoxSchemes.SelectedItem as XpColorScheme;
                if (SelectedScheme == null)
                {
                    e.Cancel = true;
                    MessageBox.Show(this, "Please select a color scheme.", "Selection Required",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            base.OnFormClosing(e);
        }

        private void listBoxSchemes_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxSchemes.SelectedItem != null)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}

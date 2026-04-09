using libmsstyle;
using msstyleEditor.PropView;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace msstyleEditor.Dialogs
{
    public partial class PropertyViewWindow : ToolWindow
    {
        private VisualStyle m_style;
        private StyleClass m_class;
        private StylePart m_part;

        private StyleState m_state;
        private StyleProperty m_prop;

        private PropertyViewMode m_viewMode;

        public delegate void PropertyChangedHandler(StyleProperty prop);
        public event PropertyChangedHandler OnPropertyAdded;
        public event PropertyChangedHandler OnPropertyRemoved;

        public PropertyViewWindow()
        {
            InitializeComponent();
            propertyView.PropertyValueChanged += (s, e) => propertyView.Refresh();
        }

        public void SetAnimation(AnimationTypeDescriptor anim)
        {
            m_viewMode = PropertyViewMode.AnimationMode;
            newPropertyToolStripMenuItem.Enabled = false;
            deleteToolStripMenuItem.Enabled = false;

            propertyView.SelectedObject = anim;
        }

        public void SetTimingFunction(TimingFunction timing)
        {
            m_viewMode = PropertyViewMode.TimingFunction;
            newPropertyToolStripMenuItem.Enabled = false;
            deleteToolStripMenuItem.Enabled = false;

            propertyView.SelectedObject = timing;
        }

        public void SetStylePart(VisualStyle style, StyleClass cls, StylePart part)
        {
            m_viewMode = PropertyViewMode.ClassMode;
            newPropertyToolStripMenuItem.Enabled = true;
            deleteToolStripMenuItem.Enabled = true;

            m_style = style;
            m_class = cls;
            m_part = part;

            propertyView.SelectedObject = (part != null)
                ? new TypeDescriptor(part, style)
                : null;
        }

        public void ShowPropertyAddDialog()
        {
            OnPropertyAdd(this, null);
        }

        public void RemoveSelectedProperty()
        {
            OnPropertyRemove(this, null);
        }


        private void OnPropertyAdd(object sender, EventArgs e)
        {
            if (m_style != null && m_class != null)
            {
                if (m_class.ClassName == "animations")
                {

                }
            }
            if(m_style == null 
                || m_class == null 
                || m_part == null 
                || m_state == null)
            {
                MessageBox.Show("Select a state or property first!", "Add Property", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var dlg = new PropertyDialog();
            dlg.StartPosition = FormStartPosition.CenterParent;
            dlg.Text = "Add Property to " + m_state.StateName;
            var newProp = dlg.ShowDialog();
            if (newProp != null)
            {
                // Add the prop where requested. TODO: test this
                newProp.Header.classID = m_class.ClassId;
                newProp.Header.partID = m_part.PartId;
                newProp.Header.stateID = m_state.StateId;

                m_state.Properties.Add(newProp);
                if(OnPropertyAdded != null)
                {
                    OnPropertyAdded(newProp);
                }
                propertyView.Refresh();
            }
        }

        private void OnPropertyRemove(object sender, EventArgs e)
        {
            if (m_viewMode != PropertyViewMode.ClassMode)
                return;
            if (m_state == null ||
                m_prop == null)
            {
                return;
            }

            m_state.Properties.Remove(m_prop);
            if (OnPropertyRemoved != null)
            {
                OnPropertyRemoved(m_prop);
            }
            propertyView.Refresh();
        }

        private void OnRecolorColors(object sender, EventArgs e)
        {
            if (m_part == null)
            {
                MessageBox.Show("Select a part first!", "Recolor", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var colorDialog = new ColorDialog())
            {
                colorDialog.FullOpen = true;
                if (colorDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var target = colorDialog.Color;
                double targetHue = RgbToHsvHue(target.R / 255.0, target.G / 255.0, target.B / 255.0);

                int count = 0;
                foreach (var state in m_part.States)
                {
                    foreach (var prop in state.Value.Properties)
                    {
                        if (prop.Header.typeID == (int)IDENTIFIER.COLOR)
                        {
                            var c = (Color)prop.GetValue();
                            var recolored = RecolorColor(c, targetHue);
                            prop.SetValue(recolored);
                            count++;
                        }
                    }
                }

                propertyView.Refresh();
                MessageBox.Show($"Recolored {count} color properties.", "Recolor", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static double RgbToHsvHue(double r, double g, double b)
        {
            double min = Math.Min(r, Math.Min(g, b));
            double max = Math.Max(r, Math.Max(g, b));
            double delta = max - min;
            if (delta < 0.00001) return 0;
            double h;
            if (r >= max)
                h = (g - b) / delta;
            else if (g >= max)
                h = 2.0 + (b - r) / delta;
            else
                h = 4.0 + (r - g) / delta;
            h *= 60.0;
            if (h < 0.0) h += 360.0;
            return h;
        }

        private static Color RecolorColor(Color c, double targetHue)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double min = Math.Min(r, Math.Min(g, b));
            double max = Math.Max(r, Math.Max(g, b));
            double delta = max - min;
            double s = max > 0.0 ? delta / max : 0;
            double v = max;

            // Convert target hue + original s,v back to RGB
            if (s <= 0.0)
                return c; // achromatic, keep as-is

            double hh = targetHue >= 360.0 ? 0.0 : targetHue;
            hh /= 60.0;
            int i = (int)hh;
            double ff = hh - i;
            double p = v * (1.0 - s);
            double q = v * (1.0 - (s * ff));
            double t = v * (1.0 - (s * (1.0 - ff)));

            double nr, ng, nb;
            switch (i)
            {
                case 0: nr = v; ng = t; nb = p; break;
                case 1: nr = q; ng = v; nb = p; break;
                case 2: nr = p; ng = v; nb = t; break;
                case 3: nr = p; ng = q; nb = v; break;
                case 4: nr = t; ng = p; nb = v; break;
                default: nr = v; ng = p; nb = q; break;
            }

            return Color.FromArgb(c.A,
                Math.Min(255, Math.Max(0, (int)(nr * 255))),
                Math.Min(255, Math.Max(0, (int)(ng * 255))),
                Math.Min(255, Math.Max(0, (int)(nb * 255))));
        }

        private void OnPropertySelected(object sender, SelectedGridItemChangedEventArgs e)
        {
            if (m_viewMode != PropertyViewMode.ClassMode)
                return;
            const int CTX_ADD = 0;
            const int CTX_REM = 1;
            const int CTX_RECOLOR = 2;

            bool haveSelection = e.NewSelection != null;
            propViewContextMenu.Items[CTX_ADD].Enabled = haveSelection;
            propViewContextMenu.Items[CTX_REM].Enabled = haveSelection;
            propViewContextMenu.Items[CTX_RECOLOR].Enabled = m_part != null;
            if (!haveSelection)
                return;

            var propDesc = e.NewSelection.PropertyDescriptor as StylePropertyDescriptor;
            if (propDesc != null)
            {
                propViewContextMenu.Items[CTX_ADD].Text = "Add Property to [" + propDesc.Category + "]";
                propViewContextMenu.Items[CTX_REM].Text = "Remove " + propDesc.Name;

                m_state = propDesc.StyleState;
                m_prop = propDesc.StyleProperty;
                return;
            }

            var dummyDesc = e.NewSelection.PropertyDescriptor as PlaceHolderPropertyDescriptor;
            if (dummyDesc != null)
            {
                propViewContextMenu.Items[CTX_ADD].Enabled = true;
                propViewContextMenu.Items[CTX_ADD].Text = "Add Property to [" + dummyDesc.Category + "]";
                propViewContextMenu.Items[CTX_REM].Enabled = false;
                propViewContextMenu.Items[CTX_REM].Text = "Remove";
                m_state = dummyDesc.StyleState;
                return;
            }

            if (e.NewSelection.GridItemType == GridItemType.Category)
            {
                OnPropertySelected(sender, new SelectedGridItemChangedEventArgs(null, e.NewSelection.GridItems[0])); // select child
                return;
            }

            if (e.NewSelection.GridItemType == GridItemType.Property)
            {
                OnPropertySelected(sender, new SelectedGridItemChangedEventArgs(null, e.NewSelection.Parent)); // select parent
                return;
            }
        }
    }

    public enum PropertyViewMode
    {
        ClassMode,
        TimingFunction,
        AnimationMode,
    }
}

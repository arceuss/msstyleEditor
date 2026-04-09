using libmsstyle;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace msstyleEditor.PropView
{
    internal sealed class FontResourceEditor : UITypeEditor
    {
        private const string DefaultFontResourceString = "Segoe UI, 9, Quality:ClearType";
        private const int DefaultFontResourceStartId = 501;

        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            return UITypeEditorEditStyle.DropDown;
        }

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            var propDesc = context?.PropertyDescriptor as StylePropertyDescriptor;
            if (propDesc == null || propDesc.Style == null)
            {
                return value;
            }

            int currentResourceId = value is int id ? id : 0;
            var availableFonts = StringTable.FilterFonts(propDesc.Style.PreferredStringTable ?? new Dictionary<int, string>());

            var editorService = provider?.GetService(typeof(IWindowsFormsEditorService)) as IWindowsFormsEditorService;
            if (editorService != null)
            {
                using (var dropDown = new FontResourceDropDown(editorService, currentResourceId, availableFonts))
                {
                    editorService.DropDownControl(dropDown);

                    if (dropDown.Action == FontResourceDropDownAction.AssignExisting &&
                        dropDown.SelectedResourceId > 0)
                    {
                        context.PropertyDescriptor.SetValue(context.Instance, dropDown.SelectedResourceId);
                        context.OnComponentChanged();
                        return dropDown.SelectedResourceId;
                    }

                    if (dropDown.Action == FontResourceDropDownAction.AddNewResource)
                    {
                        int sourceResourceId = dropDown.SelectedResourceId > 0
                            ? dropDown.SelectedResourceId
                            : currentResourceId;

                        int newResourceId = FindNextAvailableResourceId(propDesc.Style);
                        propDesc.Style.PreferredStringTable[newResourceId] = ResolveSeedFontString(propDesc.Style, sourceResourceId);

                        context.PropertyDescriptor.SetValue(context.Instance, newResourceId);
                        context.OnComponentChanged();
                        return newResourceId;
                    }

                    if (dropDown.Action != FontResourceDropDownAction.EditResource)
                    {
                        return value;
                    }

                    currentResourceId = dropDown.EditTargetResourceId;
                }
            }

            return EditWithDialog(context, value, propDesc, currentResourceId, availableFonts);
        }

        private static object EditWithDialog(
            ITypeDescriptorContext context,
            object originalValue,
            StylePropertyDescriptor propDesc,
            int currentResourceId,
            Dictionary<int, string> availableFonts)
        {
            string currentFontString = null;
            if (currentResourceId > 0)
            {
                propDesc.Style.PreferredStringTable.TryGetValue(currentResourceId, out currentFontString);
            }

            using (var dialog = new FontResourceDialog(currentResourceId, currentFontString, availableFonts))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return originalValue;
                }

                int resolvedResourceId = dialog.UseExistingResource
                    ? dialog.SelectedResourceId
                    : ResolveTargetResourceId(propDesc.Style, currentResourceId);

                if (!dialog.UseExistingResource)
                {
                    propDesc.Style.PreferredStringTable[resolvedResourceId] = dialog.EditedFontString;
                }

                context.PropertyDescriptor.SetValue(context.Instance, resolvedResourceId);
                context.OnComponentChanged();
                return resolvedResourceId;
            }
        }

        private static int ResolveTargetResourceId(VisualStyle style, int currentResourceId)
        {
            if (currentResourceId > 0)
            {
                return currentResourceId;
            }

            return FindNextAvailableResourceId(style);
        }

        private static int FindNextAvailableResourceId(VisualStyle style)
        {
            if (style.PreferredStringTable == null || style.PreferredStringTable.Count == 0)
            {
                return DefaultFontResourceStartId;
            }

            var assignedIds = new HashSet<int>(style.PreferredStringTable.Keys.Where(id => id > 0));
            var existingFontIds = StringTable.FilterFonts(style.PreferredStringTable)
                .Keys
                .Where(id => id > 0)
                .ToList();

            int nextId = existingFontIds.Count > 0
                ? existingFontIds.Max() + 1
                : Math.Max(DefaultFontResourceStartId, assignedIds.Max() + 1);

            while (assignedIds.Contains(nextId))
            {
                nextId++;
            }

            return nextId;
        }

        private static string ResolveSeedFontString(VisualStyle style, int sourceResourceId)
        {
            string seedFontString;
            if (sourceResourceId > 0 &&
                style.PreferredStringTable != null &&
                style.PreferredStringTable.TryGetValue(sourceResourceId, out seedFontString) &&
                !String.IsNullOrWhiteSpace(seedFontString))
            {
                return seedFontString;
            }

            return DefaultFontResourceString;
        }
    }

    internal enum FontResourceDropDownAction
    {
        None,
        AssignExisting,
        AddNewResource,
        EditResource,
    }

    internal sealed class ExistingFontEntry
    {
        public int Id { get; set; }
        public string FontString { get; set; }

        public override string ToString()
        {
            return $"{Id} - {FontString}";
        }
    }

    internal sealed class FontResourceDropDown : UserControl
    {
        private readonly IWindowsFormsEditorService m_editorService;
        private readonly int m_currentResourceId;
        private readonly ListBox m_listBox = new ListBox();
        private readonly Button m_assignButton = new Button();
        private readonly Button m_addButton = new Button();
        private readonly Button m_editButton = new Button();

        public FontResourceDropDownAction Action { get; private set; }
        public int SelectedResourceId => (m_listBox.SelectedItem as ExistingFontEntry)?.Id ?? 0;
        public int EditTargetResourceId { get; private set; }

        public FontResourceDropDown(IWindowsFormsEditorService editorService, int currentResourceId, Dictionary<int, string> availableFonts)
        {
            m_editorService = editorService;
            m_currentResourceId = currentResourceId;

            InitializeControl();
            LoadEntries(availableFonts);
        }

        private void InitializeControl()
        {
            Size = new Size(440, 260);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(6),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var label = new Label
            {
                AutoSize = true,
                Text = "Select an existing font resource, add a new one, or open the editor.",
                Margin = new Padding(0, 0, 0, 6),
            };
            root.Controls.Add(label, 0, 0);

            m_listBox.Dock = DockStyle.Fill;
            m_listBox.IntegralHeight = false;
            m_listBox.DoubleClick += OnAssignExisting;
            m_listBox.KeyDown += OnListBoxKeyDown;
            root.Controls.Add(m_listBox, 0, 1);

            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 0),
            };
            root.Controls.Add(buttonPanel, 0, 2);

            m_assignButton.Text = "Use Selected";
            m_assignButton.AutoSize = true;
            m_assignButton.Click += OnAssignExisting;
            buttonPanel.Controls.Add(m_assignButton);

            m_addButton.Text = "Add New";
            m_addButton.AutoSize = true;
            m_addButton.Click += OnAddNewResource;
            buttonPanel.Controls.Add(m_addButton);

            m_editButton.AutoSize = true;
            m_editButton.Click += OnEditResource;
            buttonPanel.Controls.Add(m_editButton);
        }

        private void LoadEntries(Dictionary<int, string> availableFonts)
        {
            foreach (var fontEntry in availableFonts
                .OrderBy(font => font.Key)
                .Select(font => new ExistingFontEntry { Id = font.Key, FontString = font.Value }))
            {
                m_listBox.Items.Add(fontEntry);
            }

            int index = -1;
            for (int i = 0; i < m_listBox.Items.Count; i++)
            {
                if (((ExistingFontEntry)m_listBox.Items[i]).Id == m_currentResourceId)
                {
                    index = i;
                    break;
                }
            }

            if (index >= 0)
            {
                m_listBox.SelectedIndex = index;
            }
            else if (m_listBox.Items.Count > 0)
            {
                m_listBox.SelectedIndex = 0;
            }

            bool hasEntries = m_listBox.Items.Count > 0;
            m_assignButton.Enabled = hasEntries;
            m_editButton.Text = hasEntries ? "Edit Selected..." : "Create Font...";
        }

        private void OnAssignExisting(object sender, EventArgs e)
        {
            if (SelectedResourceId <= 0)
            {
                return;
            }

            Action = FontResourceDropDownAction.AssignExisting;
            m_editorService.CloseDropDown();
        }

        private void OnEditResource(object sender, EventArgs e)
        {
            Action = FontResourceDropDownAction.EditResource;
            EditTargetResourceId = SelectedResourceId > 0 ? SelectedResourceId : m_currentResourceId;
            m_editorService.CloseDropDown();
        }

        private void OnAddNewResource(object sender, EventArgs e)
        {
            Action = FontResourceDropDownAction.AddNewResource;
            m_editorService.CloseDropDown();
        }

        private void OnListBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            e.Handled = true;
            OnAssignExisting(sender, EventArgs.Empty);
        }
    }

    internal sealed class FontResourceDialog : Form
    {
        private sealed class FontResourceDefinition
        {
            private const string NoQualityDisplayName = "(none)";

            private static readonly Dictionary<string, byte> s_qualityByName = new Dictionary<string, byte>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = 0,
                ["draft"] = 1,
                ["proof"] = 2,
                ["nonantialiased"] = 3,
                ["non-antialiased"] = 3,
                ["antialiased"] = 4,
                ["anti-aliased"] = 4,
                ["cleartype"] = 5,
                ["cleartypenatural"] = 6,
                ["cleartype-natural"] = 6,
            };

            private static readonly Dictionary<byte, string> s_nameByQuality = new Dictionary<byte, string>
            {
                [3] = "NonAntialiased",
                [4] = "Antialiased",
                [5] = "ClearType",
                [6] = "ClearType-Natural",
            };

            public string FamilyName { get; set; } = "Segoe UI";
            public int SizeInPoints { get; set; } = 9;
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public bool Underline { get; set; }
            public byte? Quality { get; set; } = 5;

            public static IReadOnlyList<string> QualityNames => new[]
            {
                NoQualityDisplayName,
                "NonAntialiased",
                "Antialiased",
                "ClearType",
                "ClearType-Natural",
            };

            public static FontResourceDefinition Parse(string fontString)
            {
                var definition = new FontResourceDefinition();
                if (String.IsNullOrWhiteSpace(fontString))
                {
                    return definition;
                }

                var parts = fontString.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim())
                    .ToList();

                if (parts.Count > 0)
                {
                    definition.FamilyName = parts[0];
                }

                if (parts.Count > 1 && Int32.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var size))
                {
                    definition.SizeInPoints = Math.Max(size, 1);
                }

                for (int i = 2; i < parts.Count; i++)
                {
                    string token = parts[i];
                    if (token.Equals("bold", StringComparison.OrdinalIgnoreCase))
                    {
                        definition.Bold = true;
                    }
                    else if (token.Equals("italic", StringComparison.OrdinalIgnoreCase))
                    {
                        definition.Italic = true;
                    }
                    else if (token.Equals("underline", StringComparison.OrdinalIgnoreCase))
                    {
                        definition.Underline = true;
                    }
                    else if (token.StartsWith("quality:", StringComparison.OrdinalIgnoreCase))
                    {
                        definition.Quality = ParseQuality(token.Substring(8));
                    }
                }

                return definition;
            }

            public string Format()
            {
                var parts = new List<string>
                {
                    FamilyName.Trim(),
                    SizeInPoints.ToString(CultureInfo.InvariantCulture)
                };

                if (Bold)
                {
                    parts.Add("Bold");
                }

                if (Italic)
                {
                    parts.Add("Italic");
                }

                if (Underline)
                {
                    parts.Add("Underline");
                }

                if (Quality.HasValue)
                {
                    parts.Add($"Quality:{GetQualityName(Quality.Value)}");
                }

                return String.Join(", ", parts);
            }

            public Font ToPreviewFont()
            {
                var style = FontStyle.Regular;
                if (Bold)
                {
                    style |= FontStyle.Bold;
                }

                if (Italic)
                {
                    style |= FontStyle.Italic;
                }

                if (Underline)
                {
                    style |= FontStyle.Underline;
                }

                try
                {
                    return new Font(FamilyName, SizeInPoints, style, GraphicsUnit.Point);
                }
                catch
                {
                    return new Font(SystemFonts.MessageBoxFont, style);
                }
            }

            public static string GetQualityName(byte quality)
            {
                return s_nameByQuality.TryGetValue(quality, out var name)
                    ? name
                    : quality.ToString(CultureInfo.InvariantCulture);
            }

            public static byte? ParseQuality(string value)
            {
                string trimmed = value.Trim();
                if (String.IsNullOrEmpty(trimmed) ||
                    trimmed.Equals(NoQualityDisplayName, StringComparison.OrdinalIgnoreCase) ||
                    trimmed.Equals("default", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                if (Byte.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric))
                {
                    return numeric;
                }

                string normalized = trimmed.Replace(" ", String.Empty);
                return s_qualityByName.TryGetValue(normalized, out var namedQuality)
                    ? (namedQuality == 0 ? (byte?)null : namedQuality)
                    : (byte?)5;
            }

            public static string GetQualityDisplayName(byte? quality)
            {
                return quality.HasValue
                    ? GetQualityName(quality.Value)
                    : NoQualityDisplayName;
            }
        }

        private readonly int m_currentResourceId;
        private readonly List<ExistingFontEntry> m_existingFonts;
        private readonly ComboBox m_existingFontCombo = new ComboBox();
        private readonly ComboBox m_familyCombo = new ComboBox();
        private readonly ComboBox m_styleCombo = new ComboBox();
        private readonly ComboBox m_sizeCombo = new ComboBox();
        private readonly CheckBox m_underlineCheck = new CheckBox();
        private readonly ComboBox m_qualityCombo = new ComboBox();
        private readonly Label m_previewLabel = new Label();
        private readonly Label m_resourceLabel = new Label();

        public bool UseExistingResource { get; private set; }
        public int SelectedResourceId { get; private set; }
        public string EditedFontString { get; private set; }

        public FontResourceDialog(int currentResourceId, string currentFontString, Dictionary<int, string> availableFonts)
        {
            m_currentResourceId = currentResourceId;
            m_existingFonts = availableFonts
                .OrderBy(font => font.Key)
                .Select(font => new ExistingFontEntry { Id = font.Key, FontString = font.Value })
                .ToList();

            SelectedResourceId = currentResourceId;

            InitializeForm();
            LoadFontChoices();
            LoadExistingFontEntries();
            LoadFontDefinition(FontResourceDefinition.Parse(currentFontString));
        }

        private void InitializeForm()
        {
            Text = "Edit Font";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 340);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(10),
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            m_resourceLabel.AutoSize = true;
            m_resourceLabel.Text = m_currentResourceId > 0
                ? $"Current font resource: {m_currentResourceId}"
                : "Current font resource: new entry will be created";
            root.Controls.Add(m_resourceLabel, 0, 0);

            var existingPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 10),
            };
            existingPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            existingPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            existingPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            root.Controls.Add(existingPanel, 0, 1);

            var existingLabel = new Label
            {
                Text = "Existing font resource:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 8, 0),
            };
            existingPanel.Controls.Add(existingLabel, 0, 0);

            m_existingFontCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            m_existingFontCombo.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            m_existingFontCombo.Width = 300;
            existingPanel.Controls.Add(m_existingFontCombo, 1, 0);

            var assignExistingButton = new Button
            {
                Text = "Assign Existing",
                AutoSize = true,
                Margin = new Padding(8, 0, 0, 0),
            };
            assignExistingButton.Click += OnAssignExisting;
            existingPanel.Controls.Add(assignExistingButton, 2, 0);

            var editorPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 10),
            };
            root.Controls.Add(editorPanel, 0, 2);

            var primaryLabels = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
                Margin = new Padding(0),
            };
            primaryLabels.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            primaryLabels.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f));
            primaryLabels.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
            editorPanel.Controls.Add(primaryLabels, 0, 0);

            primaryLabels.Controls.Add(CreateFieldLabel("Font:"), 0, 0);
            primaryLabels.Controls.Add(CreateFieldLabel("Style:"), 1, 0);
            primaryLabels.Controls.Add(CreateFieldLabel("Size:"), 2, 0);

            var primaryInputs = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 0),
            };
            primaryInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            primaryInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150f));
            primaryInputs.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90f));
            editorPanel.Controls.Add(primaryInputs, 0, 1);

            m_familyCombo.Dock = DockStyle.Top;
            m_styleCombo.Dock = DockStyle.Top;
            m_sizeCombo.Dock = DockStyle.Top;
            primaryInputs.Controls.Add(m_familyCombo, 0, 0);
            primaryInputs.Controls.Add(m_styleCombo, 1, 0);
            primaryInputs.Controls.Add(m_sizeCombo, 2, 0);

            var optionsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 10, 0, 0),
            };
            editorPanel.Controls.Add(optionsPanel, 0, 2);

            var qualityLabel = CreateFieldLabel("Quality:");
            qualityLabel.Margin = new Padding(0, 6, 8, 0);
            optionsPanel.Controls.Add(qualityLabel);

            m_qualityCombo.Width = 150;
            optionsPanel.Controls.Add(m_qualityCombo);

            m_underlineCheck.Text = "Underline";
            m_underlineCheck.AutoSize = true;
            m_underlineCheck.Margin = new Padding(16, 6, 0, 0);
            m_underlineCheck.CheckedChanged += OnEditorValueChanged;
            optionsPanel.Controls.Add(m_underlineCheck);

            var previewGroup = new GroupBox
            {
                Text = "Preview",
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
            };
            root.Controls.Add(previewGroup, 0, 3);

            m_previewLabel.Dock = DockStyle.Fill;
            m_previewLabel.TextAlign = ContentAlignment.MiddleCenter;
            m_previewLabel.Text = "AaBbYyZz";
            m_previewLabel.BackColor = SystemColors.Window;
            previewGroup.Controls.Add(m_previewLabel);

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 0),
            };
            root.Controls.Add(buttonPanel, 0, 3);

            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            var okButton = new Button { Text = "OK", AutoSize = true };
            okButton.Click += OnAccept;
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            m_familyCombo.DropDownStyle = ComboBoxStyle.DropDown;
            m_familyCombo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            m_familyCombo.AutoCompleteSource = AutoCompleteSource.ListItems;
            m_familyCombo.TextChanged += OnEditorValueChanged;

            m_styleCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            m_styleCombo.Items.AddRange(new object[]
            {
                "Regular",
                "Bold",
                "Italic",
                "Bold Italic",
            });
            m_styleCombo.SelectedIndexChanged += OnEditorValueChanged;

            m_sizeCombo.DropDownStyle = ComboBoxStyle.DropDown;
            m_sizeCombo.Items.AddRange(new object[]
            {
                "6", "7", "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"
            });
            m_sizeCombo.TextChanged += OnEditorValueChanged;

            m_qualityCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (var qualityName in FontResourceDefinition.QualityNames)
            {
                m_qualityCombo.Items.Add(qualityName);
            }
            m_qualityCombo.SelectedIndexChanged += OnEditorValueChanged;
        }

        private static Label CreateFieldLabel(string labelText)
        {
            return new Label
            {
                Text = labelText,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 0, 0),
            };
        }

        private void LoadFontChoices()
        {
            var families = FontFamily.Families
                .Select(family => family.Name)
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToArray();

            m_familyCombo.Items.AddRange(families);
        }

        private void LoadExistingFontEntries()
        {
            m_existingFontCombo.Items.Clear();
            foreach (var fontEntry in m_existingFonts)
            {
                m_existingFontCombo.Items.Add(fontEntry);
            }

            if (m_existingFontCombo.Items.Count == 0)
            {
                m_existingFontCombo.Enabled = false;
                return;
            }

            int index = m_existingFonts.FindIndex(font => font.Id == m_currentResourceId);
            m_existingFontCombo.SelectedIndex = index >= 0 ? index : 0;
        }

        private void LoadFontDefinition(FontResourceDefinition definition)
        {
            m_familyCombo.Text = definition.FamilyName;
            m_styleCombo.SelectedIndex = GetStyleIndex(definition.Bold, definition.Italic);
            m_sizeCombo.Text = definition.SizeInPoints.ToString(CultureInfo.InvariantCulture);
            m_underlineCheck.Checked = definition.Underline;
            m_qualityCombo.SelectedItem = FontResourceDefinition.GetQualityDisplayName(definition.Quality);

            if (m_qualityCombo.SelectedIndex < 0 && m_qualityCombo.Items.Count > 0)
            {
                m_qualityCombo.SelectedIndex = 0;
            }

            UpdatePreview();
        }

        private static int GetStyleIndex(bool bold, bool italic)
        {
            if (bold && italic)
            {
                return 3;
            }

            if (bold)
            {
                return 1;
            }

            if (italic)
            {
                return 2;
            }

            return 0;
        }

        private FontResourceDefinition BuildDefinitionFromEditor()
        {
            if (!Int32.TryParse(m_sizeCombo.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var size) || size <= 0)
            {
                throw new InvalidOperationException("Font size must be a positive integer.");
            }

            string family = m_familyCombo.Text.Trim();
            if (String.IsNullOrEmpty(family))
            {
                throw new InvalidOperationException("Font name cannot be empty.");
            }

            return new FontResourceDefinition
            {
                FamilyName = family,
                SizeInPoints = size,
                Bold = m_styleCombo.SelectedIndex == 1 || m_styleCombo.SelectedIndex == 3,
                Italic = m_styleCombo.SelectedIndex == 2 || m_styleCombo.SelectedIndex == 3,
                Underline = m_underlineCheck.Checked,
                Quality = FontResourceDefinition.ParseQuality(Convert.ToString(m_qualityCombo.SelectedItem ?? "(none)", CultureInfo.InvariantCulture)),
            };
        }

        private void UpdatePreview()
        {
            try
            {
                using (var previewFont = BuildDefinitionFromEditor().ToPreviewFont())
                {
                    m_previewLabel.Font = (Font)previewFont.Clone();
                }
            }
            catch
            {
                m_previewLabel.Font = SystemFonts.MessageBoxFont;
            }
        }

        private void OnEditorValueChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void OnAssignExisting(object sender, EventArgs e)
        {
            if (!(m_existingFontCombo.SelectedItem is ExistingFontEntry entry))
            {
                return;
            }

            UseExistingResource = true;
            SelectedResourceId = entry.Id;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void OnAccept(object sender, EventArgs e)
        {
            try
            {
                var definition = BuildDefinitionFromEditor();
                UseExistingResource = false;
                SelectedResourceId = m_currentResourceId;
                EditedFontString = definition.Format();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(this, ex.Message, "Edit Font", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Greenshot.Base.Controls;
using Newtonsoft.Json;
using Greenshot.Macro;

namespace Greenshot.Forms
{
    public class MacroEditorForm : GreenshotForm
    {
        private MacroScript _script;
        private ListBox _lbSteps;
        private ListBox _lbElements;
        private Panel _pnlProps;
        private TextBox _txtPreview;
        private FlowLayoutPanel _pnlActionButtons;

        // Current selection tracking
        private int _selectedStepIndex = -1;
        private int _selectedElementIndex = -1;

        public MacroEditorForm()
        {
            InitializeComponent();
            NewScript();
        }

        private void InitializeComponent()
        {
            this.Text = "Greenshot Macro Editor";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;

            var mainTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3 };
            mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // Left: Steps
            mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); // Middle: Props
            mainTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); // Right: Preview

            // --- Left Panel: Steps ---
            var pnlLeft = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            var tableLeft = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            tableLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            tableLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tableLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

            var lblSteps = new Label { Text = "Workflow Steps:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            _lbSteps = new ListBox { Dock = DockStyle.Fill };
            _lbSteps.SelectedIndexChanged += (s, e) => {
                _selectedStepIndex = _lbSteps.SelectedIndex;
                _selectedElementIndex = -1;
                LoadStepProperties();
            };
            
            var pnlStepButtons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            var btnAddStep = new Button { Text = "Add Step", AutoSize = true };
            btnAddStep.Click += (s, e) => AddStepDialog();
            var btnRemoveStep = new Button { Text = "Remove Step", AutoSize = true };
            btnRemoveStep.Click += (s, e) => RemoveStep();
            pnlStepButtons.Controls.AddRange(new Control[] { btnAddStep, btnRemoveStep });
            
            tableLeft.Controls.Add(lblSteps, 0, 0);
            tableLeft.Controls.Add(_lbSteps, 0, 1);
            tableLeft.Controls.Add(pnlStepButtons, 0, 2);
            pnlLeft.Controls.Add(tableLeft);
            mainTable.Controls.Add(pnlLeft, 0, 0);

            // --- Middle Panel: Properties ---
            _pnlProps = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            mainTable.Controls.Add(_pnlProps, 1, 0);

            // --- Right Panel: Preview ---
            var pnlRight = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            var tableRight = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            tableRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            tableRight.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            tableRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

            var lblPreview = new Label { Text = "JSON Preview:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            _txtPreview = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Font = new Font("Consolas", 10) };
            
            _pnlActionButtons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            var btnLoad = new Button { Text = "Load JSON", AutoSize = true };
            btnLoad.Click += (s, e) => LoadJson();
            var btnSave = new Button { Text = "Save JSON", AutoSize = true };
            btnSave.Click += (s, e) => SaveJson();
            var btnRun = new Button { Text = "Run Macro", AutoSize = true };
            btnRun.Click += (s, e) => RunMacro();
            _pnlActionButtons.Controls.AddRange(new Control[] { btnLoad, btnSave, btnRun });
            
            tableRight.Controls.Add(lblPreview, 0, 0);
            tableRight.Controls.Add(_txtPreview, 0, 1);
            tableRight.Controls.Add(_pnlActionButtons, 0, 2);
            pnlRight.Controls.Add(tableRight);
            mainTable.Controls.Add(pnlRight, 2, 0);

            this.Controls.Add(mainTable);
        }

        private void NewScript()
        {
            _script = new MacroScript { Version = "1.0", Workflow = new List<MacroStep>() };
            UpdateStepList();
        }

        private void UpdateStepList()
        {
            int prevIndex = _lbSteps.SelectedIndex;
            _lbSteps.Items.Clear();
            foreach (var step in _script.Workflow)
            {
                _lbSteps.Items.Add($"{_lbSteps.Items.Count + 1}. {step.Step} ({step.Type})");
            }
            if (prevIndex >= 0 && prevIndex < _lbSteps.Items.Count) _lbSteps.SelectedIndex = prevIndex;
            UpdatePreview();
        }

        private void LoadStepProperties()
        {
            _pnlProps.Controls.Clear();
            if (_selectedStepIndex < 0) return;

            var step = _script.Workflow[_selectedStepIndex];
            
            // Step Type
            var grpType = CreateGroupBox("Step Configuration", 0);
            var flowType = CreateFlow(grpType);
            
            flowType.Controls.Add(new Label { Text = "Step Type:", Width = 70 });
            var cbType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cbType.Items.AddRange(new string[] { "capture", "annotate", "export" });
            cbType.SelectedItem = step.Step;
            cbType.SelectedIndexChanged += (s, e) => {
                string newType = cbType.SelectedItem.ToString();
                if (step.Step != newType) {
                    step.Step = newType;
                    if (newType == "capture") { step.Type = "fullscreen"; step.Elements = null; step.Destinations = null; }
                    else if (newType == "annotate") { step.Type = null; step.Elements = new List<MacroElement>(); step.Destinations = null; }
                    else if (newType == "export") { step.Type = null; step.Elements = null; step.Destinations = new List<MacroDestination>(); }
                    UpdateStepList();
                    LoadStepProperties();
                }
            };
            flowType.Controls.Add(cbType);

            if (step.Step == "capture") BuildCaptureProps(step);
            else if (step.Step == "annotate") BuildAnnotateProps(step);
            else if (step.Step == "export") BuildExportProps(step);

            // Add the type selector LAST so it docks to the very TOP
            _pnlProps.Controls.Add(grpType);
        }

        private void BuildCaptureProps(MacroStep step)
        {
            var grp = CreateGroupBox("Capture Settings", 80);
            var flow = CreateFlow(grp);

            // Mode
            flow.Controls.Add(new Label { Text = "Mode:", Width = 70 });
            var cbMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cbMode.Items.AddRange(new string[] { "fullscreen", "region", "file" });
            cbMode.SelectedItem = step.Type ?? "fullscreen";
            cbMode.SelectedIndexChanged += (s, e) => {
                step.Type = cbMode.SelectedItem.ToString();
                UpdateStepList();
                LoadStepProperties();
            };
            flow.Controls.Add(cbMode);

            if (step.Type == "region")
            {
                if (step.Area == null) step.Area = new MacroArea { Width = 800, Height = 600 };
                flow.Controls.Add(CreateNumericRow("X:", step.Area.X, v => { step.Area.X = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("Y:", step.Area.Y, v => { step.Area.Y = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("Width:", step.Area.Width, v => { step.Area.Width = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("Height:", step.Area.Height, v => { step.Area.Height = (int)v; UpdatePreview(); }));
            }
            else if (step.Type == "file")
            {
                var row = new FlowLayoutPanel { Width = 350, Height = 30 };
                row.Controls.Add(new Label { Text = "Path:", Width = 50 });
                var txtPath = new TextBox { Text = step.Path, Width = 200 };
                txtPath.TextChanged += (s, e) => { step.Path = txtPath.Text; UpdatePreview(); };
                var btnBrowse = new Button { Text = "...", Width = 30 };
                btnBrowse.Click += (s, e) => {
                    using (var ofd = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.bmp" })
                        if (ofd.ShowDialog() == DialogResult.OK) txtPath.Text = ofd.FileName;
                };
                row.Controls.AddRange(new Control[] { txtPath, btnBrowse });
                flow.Controls.Add(row);
            }

            // Autocrop
            var chkAuto = new CheckBox { Text = "Enable Autocrop", Checked = step.AutoCrop ?? false, Width = 150 };
            chkAuto.CheckedChanged += (s, e) => {
                step.AutoCrop = chkAuto.Checked;
                if (!chkAuto.Checked) step.AutoCropDifference = null;
                else if (!step.AutoCropDifference.HasValue) step.AutoCropDifference = 10;
                LoadStepProperties();
                UpdatePreview();
            };
            flow.Controls.Add(chkAuto);

            if (step.AutoCrop == true)
            {
                flow.Controls.Add(CreateNumericRow("Diff:", step.AutoCropDifference ?? 10, v => { step.AutoCropDifference = (int)v; UpdatePreview(); }));
            }

            _pnlProps.Controls.Add(grp);
        }

        private void BuildAnnotateProps(MacroStep step)
        {
            var grp = CreateGroupBox("Annotations", 80);
            var flow = CreateFlow(grp);

            _lbElements = new ListBox { Width = 350, Height = 100 };
            foreach (var el in step.Elements) _lbElements.Items.Add($"{_lbElements.Items.Count + 1}. {el.Type}");
            _lbElements.SelectedIndexChanged += (s, e) => {
                _selectedElementIndex = _lbElements.SelectedIndex;
                BuildElementEditor(step);
            };
            flow.Controls.Add(_lbElements);

            var pnlBtns = new FlowLayoutPanel { Width = 350, Height = 35 };
            var cbAdd = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100 };
            cbAdd.Items.AddRange(new string[] { "rectangle", "arrow", "text", "obfuscate" });
            cbAdd.SelectedIndex = 0;
            var btnAdd = new Button { Text = "Add", Width = 60 };
            btnAdd.Click += (s, e) => { AddElement(step, cbAdd.SelectedItem.ToString()); UpdateStepList(); LoadStepProperties(); };
            var btnRem = new Button { Text = "Remove", Width = 80 };
            btnRem.Click += (s, e) => { if (_lbElements.SelectedIndex >= 0) { step.Elements.RemoveAt(_lbElements.SelectedIndex); LoadStepProperties(); UpdatePreview(); } };
            
            pnlBtns.Controls.AddRange(new Control[] { cbAdd, btnAdd, btnRem });
            flow.Controls.Add(pnlBtns);

            _pnlProps.Controls.Add(grp);

            // Container for element specific properties
            var pnlElEditor = new Panel { Name = "pnlElEditor", Width = 380, AutoSize = true, Dock = DockStyle.Top };
            _pnlProps.Controls.Add(pnlElEditor);
            
            if (_selectedElementIndex >= 0) BuildElementEditor(step);
        }

        private void BuildElementEditor(MacroStep step)
        {
            var pnl = _pnlProps.Controls.Find("pnlElEditor", false).FirstOrDefault() as Panel;
            if (pnl == null) return;
            pnl.Controls.Clear();

            if (_selectedElementIndex < 0 || _selectedElementIndex >= step.Elements.Count) return;

            var el = step.Elements[_selectedElementIndex];
            var grp = CreateGroupBox($"Element: {el.Type}", 200);
            var flow = CreateFlow(grp);

            // Bounds / Coords
            if (el.Bounds != null) {
                flow.Controls.Add(CreateNumericRow("X:", el.Bounds.X, v => { el.Bounds.X = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("Y:", el.Bounds.Y, v => { el.Bounds.Y = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("W:", el.Bounds.W, v => { el.Bounds.W = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("H:", el.Bounds.H, v => { el.Bounds.H = (int)v; UpdatePreview(); }));
            } else if (el.From != null) {
                flow.Controls.Add(CreateNumericRow("From X:", el.From.X, v => { el.From.X = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("From Y:", el.From.Y, v => { el.From.Y = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("To X:", el.To.X, v => { el.To.X = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("To Y:", el.To.Y, v => { el.To.Y = (int)v; UpdatePreview(); }));
            } else if (el.Position != null) {
                flow.Controls.Add(CreateNumericRow("X:", el.Position.X, v => { el.Position.X = (int)v; UpdatePreview(); }));
                flow.Controls.Add(CreateNumericRow("Y:", el.Position.Y, v => { el.Position.Y = (int)v; UpdatePreview(); }));
            }

            if (el.Content != null) {
                var row = new FlowLayoutPanel { Width = 350, Height = 30 };
                row.Controls.Add(new Label { Text = "Text:", Width = 50 });
                var txt = new TextBox { Text = el.Content, Width = 250 };
                txt.TextChanged += (s, e) => { el.Content = txt.Text; UpdatePreview(); };
                row.Controls.Add(txt);
                flow.Controls.Add(row);
            }

            // Styles
            if (el.Style == null) el.Style = new MacroStyle();
            flow.Controls.Add(new Label { Text = "Styles:", Font = new Font(this.Font, FontStyle.Bold) });
            
            flow.Controls.Add(CreateColorRow("Line Color:", el.Style.LineColor, c => { el.Style.LineColor = c; UpdatePreview(); }));
            flow.Controls.Add(CreateColorRow("Fill Color:", el.Style.FillColor, c => { el.Style.FillColor = c; UpdatePreview(); }));
            flow.Controls.Add(CreateNumericRow("Thickness:", el.Style.LineThickness ?? 3, v => { el.Style.LineThickness = (int)v; UpdatePreview(); }));
            
            var chkShadow = new CheckBox { Text = "Drop Shadow", Checked = el.Style.Shadow ?? false };
            chkShadow.CheckedChanged += (s, e) => { el.Style.Shadow = chkShadow.Checked; UpdatePreview(); };
            flow.Controls.Add(chkShadow);

            if (el.Type == "obfuscate") {
               flow.Controls.Add(CreateNumericRow("Blur:", el.Style.BlurRadius ?? 5, v => { el.Style.BlurRadius = (int)v; UpdatePreview(); }));
               flow.Controls.Add(CreateNumericRow("Pixel:", el.Style.PixelSize ?? 5, v => { el.Style.PixelSize = (int)v; UpdatePreview(); }));
            }
            if (el.Type == "text") {
               flow.Controls.Add(CreateNumericRow("Font Size:", (int)(el.Style.FontSize ?? 12), v => { el.Style.FontSize = (double)v; UpdatePreview(); }));
            }

            pnl.Controls.Add(grp);
        }

        private void BuildExportProps(MacroStep step)
        {
            var grp = CreateGroupBox("Export Destinations", 80);
            var flow = CreateFlow(grp);

            foreach (var dest in step.Destinations) {
                flow.Controls.Add(new Label { Text = $"- {dest.Type}", Width = 300 });
                if (dest.Type == "file") flow.Controls.Add(new Label { Text = $"  Path: {dest.Path}", Width = 300, Font = new Font(this.Font.FontFamily, 8) });
            }

            var btnAddFile = new Button { Text = "+ File", AutoSize = true };
            btnAddFile.Click += (s, e) => {
                using (var sfd = new SaveFileDialog { Filter = "PNG|*.png|JPG|*.jpg" })
                    if (sfd.ShowDialog() == DialogResult.OK) {
                        step.Destinations.Add(new MacroDestination { Type = "file", Path = sfd.FileName });
                        LoadStepProperties();
                        UpdatePreview();
                    }
            };
            var btnAddClip = new Button { Text = "+ Clipboard", AutoSize = true };
            btnAddClip.Click += (s, e) => { step.Destinations.Add(new MacroDestination { Type = "clipboard" }); LoadStepProperties(); UpdatePreview(); };

            var pnlBtns = new FlowLayoutPanel { Width = 350, Height = 35, AutoSize = true };
            pnlBtns.Controls.AddRange(new Control[] { btnAddFile, btnAddClip });
            flow.Controls.Add(pnlBtns);

            _pnlProps.Controls.Add(grp);
        }

        // --- Helpers ---
        private GroupBox CreateGroupBox(string text, int top) {
            return new GroupBox { Text = text, Width = 400, AutoSize = true, Dock = DockStyle.Top };
        }

        private FlowLayoutPanel CreateFlow(Control parent) {
            var f = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true, Padding = new Padding(5) };
            parent.Controls.Add(f);
            return f;
        }

        private FlowLayoutPanel CreateNumericRow(string label, int value, Action<decimal> onChanged) {
            var row = new FlowLayoutPanel { Width = 350, Height = 30 };
            row.Controls.Add(new Label { Text = label, Width = 80, TextAlign = ContentAlignment.MiddleLeft });
            var num = new NumericUpDown { Minimum = -5000, Maximum = 5000, Width = 80 };
            num.Value = value;
            num.ValueChanged += (s, e) => onChanged(num.Value);
            row.Controls.Add(num);
            return row;
        }

        private FlowLayoutPanel CreateColorRow(string label, string value, Action<string> onChanged) {
            var row = new FlowLayoutPanel { Width = 350, Height = 30 };
            row.Controls.Add(new Label { Text = label, Width = 80 });
            var txt = new TextBox { Text = value ?? "transparent", Width = 100 };
            txt.TextChanged += (s, e) => onChanged(txt.Text);
            var btn = new Button { Text = "Pick", Width = 50 };
            btn.Click += (s, e) => {
                var cd = new ColorDialog();
                if (cd.ShowDialog() == DialogResult.OK) txt.Text = ColorTranslator.ToHtml(cd.Color);
            };
            row.Controls.AddRange(new Control[] { txt, btn });
            return row;
        }

        private void AddElement(MacroStep step, string type)
        {
            var el = new MacroElement { Type = type, Style = new MacroStyle() };
            if (type == "rectangle" || type == "obfuscate") el.Bounds = new MacroBounds { X = 100, Y = 100, W = 200, H = 150 };
            else if (type == "arrow") { el.From = new MacroPoint { X = 100, Y = 100 }; el.To = new MacroPoint { X = 200, Y = 200 }; }
            else if (type == "text") { el.Position = new MacroPoint { X = 100, Y = 100 }; el.Content = "New Text"; }
            step.Elements.Add(el);
            UpdatePreview();
        }

        private void AddStepDialog()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Capture", null, (s, e) => { _script.Workflow.Add(new MacroStep { Step = "capture", Type = "fullscreen" }); UpdateStepList(); });
            menu.Items.Add("Annotate", null, (s, e) => { _script.Workflow.Add(new MacroStep { Step = "annotate", Elements = new List<MacroElement>() }); UpdateStepList(); });
            menu.Items.Add("Export", null, (s, e) => { _script.Workflow.Add(new MacroStep { Step = "export", Destinations = new List<MacroDestination>() }); UpdateStepList(); });
            menu.Show(Cursor.Position);
        }

        private void RemoveStep() { if (_selectedStepIndex >= 0) { _script.Workflow.RemoveAt(_selectedStepIndex); _selectedStepIndex = -1; UpdateStepList(); LoadStepProperties(); } }
        private void UpdatePreview() { _txtPreview.Text = JsonConvert.SerializeObject(_script, Formatting.Indented); }
        private void SaveJson() { using (var sfd = new SaveFileDialog { Filter = "JSON|*.json" }) if (sfd.ShowDialog() == DialogResult.OK) File.WriteAllText(sfd.FileName, _txtPreview.Text); }
        private void LoadJson() { using (var ofd = new OpenFileDialog { Filter = "JSON|*.json" }) if (ofd.ShowDialog() == DialogResult.OK) { _script = JsonConvert.DeserializeObject<MacroScript>(File.ReadAllText(ofd.FileName)); UpdateStepList(); LoadStepProperties(); } }
        private void RunMacro() { var p = Path.GetTempFileName() + ".json"; File.WriteAllText(p, _txtPreview.Text); MacroEngine.Run(p); }
    }
}

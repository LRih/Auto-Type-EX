using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AutoTypeEX
{
    public class AutoTypeEX : Form
    {
        //===================================================================== CONSTANTS
        private readonly Font DEFAULT_FONT = new Font("Microsoft Sans Serif", 8); // default used font
        private readonly string[] COLOR_COMMANDS = new string[] { "", "cyan:", "green:", "purple:", "red:", "white:" };
        private readonly Color[] COLORS = new Color[] { Color.Yellow, Color.SkyBlue, Color.LimeGreen, Color.MediumOrchid, Color.IndianRed, Color.WhiteSmoke };
        private const char DELIM = '\x7C';

        //===================================================================== CONTROLS
        private IContainer _components = new Container();

        private TextBox[] _txtData = new TextBox[9];
        private ComboBox[] _comColors = new ComboBox[9];
        private NumericUpDown _numMin = new NumericUpDown();
        private NumericUpDown _numMax = new NumericUpDown();
        private GroupBox _grpSpeed = new GroupBox();

        private PictureBox _picCalculator = new PictureBox();
        private PictureBox _picMinimizer = new PictureBox();
        private PictureBox _picNotepad = new PictureBox();
        private PictureBox _picResizer = new PictureBox();
        private PictureBox _picHide = new PictureBox();

        private ContextMenuStrip _contextMenu;
        private ToolStripMenuItem _menuAbout = new ToolStripMenuItem();
        private ToolStripMenuItem _menuOptions = new ToolStripMenuItem();
        private ToolStripMenuItem _menuTextColor = new ToolStripMenuItem();
        private ToolStripMenuItem _menuTextboxColor = new ToolStripMenuItem();
        private ToolStripMenuItem _menuUITextColor = new ToolStripMenuItem();
        private ToolStripMenuItem _menuEnabledColor = new ToolStripMenuItem();
        private ToolStripMenuItem _menuDisabledColor = new ToolStripMenuItem();
        private ToolStripMenuItem _menuFont = new ToolStripMenuItem();
        private ToolStripMenuItem _menuTopmost = new ToolStripMenuItem();
        private ToolStripMenuItem _menuExit = new ToolStripMenuItem();

        private NotifyIcon _notifyIcon;
        private Timer _timer;

        //===================================================================== VARIABLES
        private bool _isTyping = false;
        private bool _isEnabled = true;
        private int _currentProfile = 0; // stores current profile

        private string[][] _text = new string[9][]; // holds text for auto-typing
        private int[][] _color = new int[9][]; // holds text color

        private Font _fntCurrent = new Font("Lucida Sans Unicode", 12); // currently used font

        // default colors
        private Color _colText = Color.FromArgb(255, 255, 255);
        private Color _colTextbox = Color.FromArgb(73, 78, 73);
        private Color _colUIText = Color.FromArgb(216, 222, 211);
        private Color _colEnabled = Color.FromArgb(70, 70, 70);
        private Color _colDisabled = Color.FromArgb(90, 106, 80);

        //===================================================================== INITIALIZE
        public AutoTypeEX()
        {
            this.ClientSize = new Size(275, 350);

            for (int i = 0; i < _txtData.Length; i++)
            {
                _txtData[i] = new TextBox();
                _txtData[i].Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                _txtData[i].BorderStyle = BorderStyle.FixedSingle;
                _txtData[i].Location = new Point(35, 6);
                _txtData[i].Width = ClientSize.Width - 35 - 30;
                _txtData[i].TabIndex = i;
                this.Controls.Add(_txtData[i]);
            }
            for (int i = 0; i < _comColors.Length; i++)
            {
                _comColors[i] = new ComboBox();
                _comColors[i].Anchor = AnchorStyles.Top | AnchorStyles.Right;
                _comColors[i].DrawMode = DrawMode.OwnerDrawVariable;
                _comColors[i].DropDownStyle = ComboBoxStyle.DropDownList;
                _comColors[i].Items.AddRange(new object[] { "N", "C", "G", "P", "R", "W" });
                _comColors[i].Location = new Point(_txtData[i].Right + 3, 6);
                _comColors[i].Width = ClientSize.Width - _txtData[i].Right - 3 + 10;
                _comColors[i].TabStop = false;
                _comColors[i].DrawItem += comColor_DrawItem;
                _comColors[i].SelectedIndex = 0;
                _comColors[i].SelectedIndexChanged += comColor_SelectedIndexChanged;
                this.Controls.Add(_comColors[i]);
            }

            _numMin.Size = _numMax.Size = new Size(75, 20);
            _numMin.Location = new Point(6, 20);
            _numMin.Maximum = 1000;
            _numMin.Minimum = 5;
            _numMin.Value = 10;
            _numMin.ValueChanged += numMin_ValueChanged;
            _numMax.Location = new Point(_numMin.Right + 30, _numMin.Top);
            _numMax.Maximum = 1000;
            _numMax.Minimum = 5;
            _numMax.Value = 35;
            _numMax.ValueChanged += numMax_ValueChanged;
            _grpSpeed.Left = 35;
            _grpSpeed.Size = new Size(300, 60);
            _grpSpeed.TabStop = false;
            _grpSpeed.Controls.Add(_numMin);
            _grpSpeed.Controls.Add(_numMax);
            _grpSpeed.Text = "Type Speed";
            _grpSpeed.Paint += grpSpeed_Paint;
            _picCalculator.BorderStyle = _picMinimizer.BorderStyle = _picNotepad.BorderStyle = _picResizer.BorderStyle = BorderStyle.FixedSingle;
            _picCalculator.Size = _picMinimizer.Size = _picNotepad.Size = _picResizer.Size = new Size(10, 10);
            _picCalculator.BackColor = Color.Yellow;
            _picCalculator.Location = new Point(0, 0);
            _picCalculator.Click += picCalculator_Click;
            _picMinimizer.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _picMinimizer.BackColor = Color.Red;
            _picMinimizer.Location = new Point(ClientSize.Width - 10, 0);
            _picMinimizer.MouseClick += picMinimizer_MouseClick;
            _picNotepad.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            _picNotepad.BackColor = System.Drawing.Color.Lime;
            _picNotepad.Location = new Point(0, ClientSize.Height - 10);
            _picResizer.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            _picResizer.BackColor = SystemColors.Highlight;
            _picResizer.Cursor = Cursors.SizeNWSE;
            _picResizer.Location = new Point(ClientSize.Width - 10, ClientSize.Height - 10);
            _picResizer.MouseDown += picResizer_MouseDown;
            _picHide.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _picHide.BackColor = Color.Transparent;
            _picHide.Location = new Point(ClientSize.Width - 5, 6);
            _picHide.Width = 5;
            _picHide.MouseDown += AutoType_MouseDown;
            _contextMenu = new ContextMenuStrip(_components);
            _contextMenu.Items.AddRange(new ToolStripItem[] { _menuAbout, _menuOptions, _menuTopmost, _menuExit });
            _menuAbout.Text = "About";
            _menuAbout.Click += menuAbout_Click;
            _menuOptions.DropDownItems.AddRange(new ToolStripItem[] { _menuTextColor, _menuTextboxColor, new ToolStripSeparator() });
            _menuOptions.DropDownItems.AddRange(new ToolStripItem[] { _menuUITextColor, new ToolStripSeparator() });
            _menuOptions.DropDownItems.AddRange(new ToolStripItem[] { _menuEnabledColor, _menuDisabledColor, new ToolStripSeparator() });
            _menuOptions.DropDownItems.AddRange(new ToolStripItem[] { _menuFont });
            _menuOptions.Text = "Options";
            _menuTextColor.Text = "Text Color";
            _menuTextColor.Click += menuTextColor_Click;
            _menuTextboxColor.Text = "Textbox Color";
            _menuTextboxColor.Click += menuTextboxColor_Click;
            _menuUITextColor.Text = "UI Text Color";
            _menuUITextColor.Click += menuUITextColor_Click;
            _menuEnabledColor.Text = "Enabled Color";
            _menuEnabledColor.Click += menuEnabledColor_Click;
            _menuDisabledColor.Text = "Disabled Color";
            _menuDisabledColor.Click += menuDisabledColor_Click;
            _menuFont.Text = "Font";
            _menuFont.Click += menuFont_Click;
            _menuTopmost.CheckOnClick = true;
            _menuTopmost.Text = "Always on Top";
            _menuTopmost.Click += menuTopmost_Click;
            _menuExit.Text = "Exit";
            _menuExit.Click += menuExit_Click;
            _notifyIcon = new NotifyIcon(_components);
            _notifyIcon.ContextMenuStrip = _contextMenu;
            _notifyIcon.Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("AutoTypeEX.Icon.ico"));
            _notifyIcon.Text = "Auto Type";
            _notifyIcon.Visible = true;
            _notifyIcon.MouseDoubleClick += notifyMain_MouseDoubleClick;
            _timer = new Timer(_components);
            _timer.Enabled = true;
            _timer.Interval = 10;
            _timer.Tick += timerMain_Tick;

            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("AutoTypeEX.Icon.ico"));
            this.MinimumSize = new Size(20, 20);
            this.Text = "Auto Type EX";
            this.Controls.Add(_grpSpeed);
            this.Controls.Add(_picCalculator);
            this.Controls.Add(_picMinimizer);
            this.Controls.Add(_picNotepad);
            this.Controls.Add(_picResizer);
            this.Controls.Add(_picHide);
            this.Load += AutoTypeEX_Load;
            this.MouseClick += AutoTypeEX_MouseClick;
            this.MouseDown += AutoType_MouseDown;
            this.Paint += AutoTypeEX_Paint;
            this.Resize += AutoType_Resize;

            LoadSettings();

            // arrange window
            ColorWindow();
            ArrangeWindow();
        }
        private void AutoTypeEX_Load(object sender, EventArgs e)
        {
            _picHide.BringToFront();
            _picCalculator.BringToFront();
            _picMinimizer.BringToFront();
            _picNotepad.BringToFront();
            _picResizer.BringToFront();
        }

        // arrange windows
        private void ColorWindow()
        {
            foreach (TextBox textbox in _txtData)
            {
                textbox.ForeColor = _colText;
                textbox.BackColor = _colTextbox;
            }

            // change UI colors
            foreach (ComboBox colorbox in _comColors)
            {
                colorbox.ForeColor = _colText;
                colorbox.BackColor = _colTextbox;
            }
            _grpSpeed.ForeColor = _colUIText;
            _numMin.ForeColor = _numMax.ForeColor = _colText;
            _numMin.BackColor = _numMax.BackColor = _colTextbox;
            this.BackColor = _colEnabled;

            // change menu colors
            _menuTextColor.ForeColor = _menuTextboxColor.ForeColor = _colText;
            _menuTextColor.BackColor = _menuTextboxColor.BackColor = _colTextbox;
            _menuUITextColor.ForeColor = _menuEnabledColor.ForeColor = _menuDisabledColor.ForeColor = _colUIText;
            _menuUITextColor.BackColor = _menuEnabledColor.BackColor = _colEnabled;
            _menuDisabledColor.BackColor = _colDisabled;
        }
        private void ArrangeWindow()
        {
            ArrangeControls();

            // resize form
            this.Height = _txtData[8].Bottom + 6;
            this.Invalidate();
        }
        private void ArrangeControls()
        {
            // change fonts
            foreach (TextBox textbox in _txtData) textbox.Font = _fntCurrent;

            // arrange controls
            for (int i = 1; i < _txtData.Length; i++) _txtData[i].Top = _txtData[i - 1].Bottom + 6;
            for (int i = 0; i < _comColors.Length; i++) _comColors[i].Location = new Point(_txtData[i].Right + 6, _txtData[i].Location.Y + (_txtData[i].Size.Height / 2 - _comColors[i].Size.Height / 2));
            
            // arrange ui
            _grpSpeed.Top = _txtData[8].Location.Y + _txtData[8].Size.Height + 6;
            _picHide.Height = _comColors[8].Bottom - _comColors[0].Top + 6;
        }

        private void LoadSettings()
        {
            for (int i = 0; i < 9; i++)
            {
                _text[i] = new string[3];
                _color[i] = new int[3];
            }
            
            if (!File.Exists(SavePath)) return;

            try
            {
                for (int i = 0; i < 9; i++)
                {
                    _text[i] = INI.GetValue("Text", string.Format("F{0}", i + 1), SavePath).Split('|');
                    _txtData[i].Text = _text[i][_currentProfile];
                    string[] colorData = INI.GetValue("Color", string.Format("F{0}", i + 1), SavePath).Split('|');
                    for (int profileID = 0; profileID < 3; profileID++) _color[i][profileID] = Convert.ToInt32(colorData[profileID]);
                    _comColors[i].SelectedIndex = _color[i][_currentProfile];
                }

                // load colors
                _colText = LoadColor("Text Color");
                _colTextbox = LoadColor("Textbox Color");
                _colUIText = LoadColor("UI Text Color");
                _colEnabled = LoadColor("Enabled Color");
                _colDisabled = LoadColor("Disabled Color");

                // load font
                _fntCurrent = new Font(INI.GetValue("Options", "Font Name", SavePath), Convert.ToInt32(INI.GetValue("Options", "Font Size", SavePath)));
                _numMin.Value = Convert.ToInt32(INI.GetValue("Options", "Type Speed Min", SavePath));
                _numMax.Value = Convert.ToInt32(INI.GetValue("Options", "Type Speed Max", SavePath));
                _menuTopmost.CheckState = (CheckState)Convert.ToInt32(Convert.ToInt32(INI.GetValue("Options", "Topmost", SavePath)));
            }
            catch
            {
                MessageBox.Show("Save file corrupted. Some settings not loaded");
            }
        }
        private void SaveSettings()
        {
            for (int i = 0; i < 9; i++)
            {
                _text[i][_currentProfile] = _txtData[i].Text;
                _color[i][_currentProfile] = _comColors[i].SelectedIndex;
            }
            for (int i = 0; i < 9; i++) INI.SetValue("Text", string.Format("F{0}", i + 1), string.Join("|", _text[i]), SavePath);
            for (int i = 0; i < 9; i++) INI.SetValue("Color", string.Format("F{0}", i + 1), string.Join("|", _color[i]), SavePath);
            INI.SetValue("Options", "Text Color", string.Format("{0},{1},{2}", _colText.R.ToString(), _colText.G.ToString(), _colText.B.ToString()), SavePath);
            INI.SetValue("Options", "Textbox Color", string.Format("{0},{1},{2}", _colTextbox.R.ToString(), _colTextbox.G.ToString(), _colTextbox.B.ToString()), SavePath);
            INI.SetValue("Options", "UI Text Color", string.Format("{0},{1},{2}", _colUIText.R.ToString(), _colUIText.G.ToString(), _colUIText.B.ToString()), SavePath);
            INI.SetValue("Options", "Enabled Color", string.Format("{0},{1},{2}", _colEnabled.R.ToString(), _colEnabled.G.ToString(), _colEnabled.B.ToString()), SavePath);
            INI.SetValue("Options", "Disabled Color", string.Format("{0},{1},{2}", _colDisabled.R.ToString(), _colDisabled.G.ToString(), _colDisabled.B.ToString()), SavePath);
            INI.SetValue("Options", "Font Name", _fntCurrent.Name, SavePath);
            INI.SetValue("Options", "Font Size", _fntCurrent.Size.ToString(), SavePath);
            INI.SetValue("Options", "Type Speed Min", _numMin.Value.ToString(), SavePath);
            INI.SetValue("Options", "Type Speed Max", _numMax.Value.ToString(), SavePath);
            INI.SetValue("Options", "Topmost", ((int)_menuTopmost.CheckState).ToString(), SavePath);
        }
        private Color LoadColor(string key)
        {
            string[] color = INI.GetValue("Options", key, SavePath).Split(',');
            return Color.FromArgb(Convert.ToInt32(color[0]), Convert.ToInt32(color[1]), Convert.ToInt32(color[2]));
        }

        //===================================================================== TERMINATE
        protected override void Dispose(bool disposing)
        {
            if (disposing && _components != null) _components.Dispose();
            base.Dispose(disposing);
        }

        //===================================================================== FUNCTIONS
        private void AskUserForColor(ref Color color)
        {
            using (ColorPicker dialog = new ColorPicker(color))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    color = dialog.SelectedColor;
            }
        }

        //===================================================================== PROPERTIES
        private string Version
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return string.Format("{0}.{1}", version.Major, version.Minor);
            }
        }
        private string SavePath
        {
            get
            {
                return string.Format(@"{0}\{1} {2}.ini", Application.StartupPath, Application.ProductName, Version);
            }
        }

        //===================================================================== EVENTS
        private void timerMain_Tick(object sender, EventArgs e)
        {
            // return if disabled or typing
            if (!_isEnabled || _isTyping) return;

            // check if changing profiles
            if (Hotkey.IsKeyDown(Keys.LControlKey) && Hotkey.IsKeyDown(Keys.MButton))
            {
                _isTyping = true;
                int oldProfile = _currentProfile;
                _currentProfile = (_currentProfile + 1) % 3;
                for (int i = 0; i < 9; i++)
                {
                    _text[i][oldProfile] = _txtData[i].Text;
                    _color[i][oldProfile] = _comColors[i].SelectedIndex;
                    _txtData[i].Text = _text[i][_currentProfile];
                    _comColors[i].SelectedIndex = _color[i][_currentProfile];
                }
                Hotkey.WaitUntilKeyUp(Keys.MButton);
                _isTyping = false;
                this.Invalidate();
            }

            // check if hotkey pressed
            for (int key = 112; key < 121; key++)
            {
                if (Hotkey.IsKeyDown((Keys)key) && Hotkey.IsKeyDown((Keys)key))
                {
                    int index = key - 112; // f index
                    _isTyping = true;

                    // type text and enter
                    Macro.TypeText(COLOR_COMMANDS[_comColors[index].SelectedIndex] + _txtData[index].Text, (int)_numMin.Value, (int)_numMax.Value);
                    Macro.TypeText("\r\n", (int)_numMin.Value, (int)_numMax.Value);
                    Hotkey.WaitUntilKeyUp((Keys)key);
                    _isTyping = false;
                    break;
                }
            }
        }

        // type speed
        private void numMin_ValueChanged(object sender, EventArgs e)
        {
            _numMax.Value = Math.Max(_numMin.Value, _numMax.Value); // ensure max is greater than min
        }
        private void numMax_ValueChanged(object sender, EventArgs e)
        {
            _numMin.Value = Math.Min(_numMin.Value, _numMax.Value); // ensure min is less than max
        }

        // minimize / restore
        private void AutoType_Resize(object sender, EventArgs e)
        {
            this.ShowInTaskbar = (this.WindowState != FormWindowState.Minimized);
        }
        private void AutoType_MouseDown(object sender, MouseEventArgs e)
        {
            Point movePos = new Point(MousePosition.X - this.Left, MousePosition.Y - this.Top);

            do
            {
                this.Location = new Point(MousePosition.X - movePos.X, MousePosition.Y - movePos.Y);
                Application.DoEvents();
            } while (Hotkey.IsKeyDown(Keys.LButton));
        }

        // enable / disable
        private void AutoTypeEX_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _isEnabled = (_isEnabled != true);
                this.BackColor = (_isEnabled ? _colEnabled : _colDisabled);
            }
        }

        // paint colors
        private void AutoTypeEX_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Brush brushText = new SolidBrush(_colUIText);

            for (int i = 0; i < _txtData.Length; i++)
            {
                int x = 12;
                int y = _txtData[i].Top + (_txtData[i].Height / 2 - DEFAULT_FONT.Height / 2);
                e.Graphics.DrawString(string.Format("F{0}", i + 1), DEFAULT_FONT, brushText, x, y);
                Rectangle rect = new Rectangle(_txtData[i].Left - 5, _txtData[i].Top, _txtData[i].Width + 10 - 1, _txtData[i].Height - 1);
                using (Brush b = new SolidBrush(COLORS[_comColors[i].SelectedIndex]))
                    e.Graphics.FillRectangle(b, rect);
            }
            e.Graphics.DrawString((_currentProfile + 1).ToString(), DEFAULT_FONT, brushText, 0, 10);

            brushText.Dispose();
        }
        private void grpSpeed_Paint(object sender, PaintEventArgs e)
        {
            TextRenderer.DrawText(e.Graphics, "to", this.Font, new Point(_numMin.Right + 10, _numMin.Top), _colUIText);
        }
        private void comColor_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            string text = ((ComboBox)sender).Items[e.Index].ToString();
            Rectangle rect = e.Bounds;
            using (Brush b = new SolidBrush(_colTextbox))
                e.Graphics.FillRectangle(b, rect.X, rect.Y - 1, rect.Width, rect.Height + 2);
            e.DrawBackground();
            using (Brush b = new SolidBrush(_colText))
                e.Graphics.DrawString(text, DEFAULT_FONT, b, rect.X, rect.Y);
        }
        private void comColor_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Invalidate();
        }

        // context menu
        private void notifyMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }
        private void menuAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Application.ProductName + " " + Version + "\r\nRic", Application.ProductName);
        }

        // color / font
        private void menuTextColor_Click(object sender, EventArgs e)
        {
            AskUserForColor(ref _colText);
            ColorWindow();
        }
        private void menuTextboxColor_Click(object sender, EventArgs e)
        {
            AskUserForColor(ref _colTextbox);
            ColorWindow();
        }
        private void menuUITextColor_Click(object sender, EventArgs e)
        {
            AskUserForColor(ref _colUIText);
            ColorWindow();
        }
        private void menuEnabledColor_Click(object sender, EventArgs e)
        {
            AskUserForColor(ref _colEnabled);
            ColorWindow();
        }
        private void menuDisabledColor_Click(object sender, EventArgs e)
        {
            AskUserForColor(ref _colDisabled);
            ColorWindow();
        }
        private void menuFont_Click(object sender, EventArgs e)
        {
            FontDialog dialog = new FontDialog();
            dialog.Font = _fntCurrent;

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _fntCurrent = dialog.Font;
                ArrangeWindow();
            }

            dialog.Dispose();
        }

        // topmost
        private void menuTopmost_Click(object sender, EventArgs e)
        {
            this.TopMost = _menuTopmost.Checked;
        }
        private void menuExit_Click(object sender, EventArgs e)
        {
            SaveSettings();
            this.Dispose();
        }

        // corner buttons
        private void picCalculator_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("calc");
        }
        private void picMinimizer_MouseClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        private void picResizer_MouseDown(object sender, MouseEventArgs e)
        {
            // position of initiation
            Point resizePos = new Point(this.Right - MousePosition.X, this.Bottom - MousePosition.Y);

            do
            {
                this.Size = new Size(MousePosition.X - this.Left + resizePos.X, MousePosition.Y - this.Top + resizePos.Y);
                this.Invalidate();
                Application.DoEvents();
            } while (Hotkey.IsKeyDown(Keys.LButton));
        }
    }
}

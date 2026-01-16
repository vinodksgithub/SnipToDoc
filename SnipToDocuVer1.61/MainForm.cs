
using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Label = System.Windows.Forms.Label;
using Font = System.Drawing.Font;

namespace ScreenCaptureUtility
{
    public partial class MainForm : Form
    {
        private PictureBox pictureBoxPreview;
        private Button btnCapture;
        private Button btnRegionCapture;
        private Button btnSaveToWord;
        private Button btnSaveClose;
        private CheckBox chkAppend;
        private CheckBox chkStayOnTop;
        private Label lblStatus;
        private Panel bottomPanel;
        private Button btnSetLocation;

        private ImageEditor _imageEditor;
        private SaveOptionsHandler _saveHandler;
        private Bitmap _currentCapture;

        public ImageEditor ImageEditor => _imageEditor;


        public MainForm()
        {
            Text = "QA Evidence Capturer v1.61";
            Size = new Size(1000, 800);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(700, 450);
            InitializeUnifiedUI();
        }

        private void InitializeUnifiedUI()
        {
            bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 150,
                BackColor = Color.FromArgb(230, 230, 230),
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(bottomPanel);

            pictureBoxPreview = new PictureBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.Fixed3D,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Gray
            };
            Controls.Add(pictureBoxPreview);

            _imageEditor = new ImageEditor(pictureBoxPreview);

            btnCapture = new Button
            {
                Text = "📸 Capture Screen",
                Location = new Point(20, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnCapture.Click += BtnCapture_Click;
            bottomPanel.Controls.Add(btnCapture);

            btnRegionCapture = new Button
            {
                Text = "📐 Region Capture",
                Location = new Point(180, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnRegionCapture.Click += BtnRegionCapture_Click;
            bottomPanel.Controls.Add(btnRegionCapture);

            btnSaveToWord = new Button
            {
                Text = "💾 Save",
                Location = new Point(340, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Enabled = false
            };
            bottomPanel.Controls.Add(btnSaveToWord);

            btnSaveClose = new Button
            {
                Text = "💾 Save & Close",
                Location = new Point(btnSaveToWord.Right + 15, btnSaveToWord.Top),
                Size = new Size(150, 40),
                Font = btnSaveToWord.Font,
                Visible = false
            };
            bottomPanel.Controls.Add(btnSaveClose);

            int chkLeft = btnSaveClose.Right + 30;

            chkAppend = new CheckBox
            {
                Text = "Append Mode (Keep Word Open)",
                Location = new Point(chkLeft, 25),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };
            bottomPanel.Controls.Add(chkAppend);

            chkStayOnTop = new CheckBox
            {
                Text = "Stay on Top",
                Location = new Point(chkLeft, 55),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };
            chkStayOnTop.CheckedChanged += (s, e) => this.TopMost = chkStayOnTop.Checked;
            bottomPanel.Controls.Add(chkStayOnTop);


            btnSetLocation = new Button
            {
                Text = "📂 Set Save Location",
                Location = new Point(chkStayOnTop.Left, chkStayOnTop.Bottom + 10),
                Size = new Size(230, 40),
                Font = new Font("Segoe UI", 9)
            };
            btnSetLocation.Click += BtnSetLocation_Click;
            bottomPanel.Controls.Add(btnSetLocation);

            lblStatus = new Label
            {
                Text = "Ready.",
                Location = new Point(20, 100),
                AutoSize = true,
                ForeColor = Color.DarkBlue
            };
            bottomPanel.Controls.Add(lblStatus);

            ToolStrip toolStrip = new ToolStrip
            {
                Dock = DockStyle.Top,
                GripStyle = ToolStripGripStyle.Hidden
            };
            Controls.Add(toolStrip);
            toolStrip.BringToFront();


            AddToolButton(toolStrip, "⬛ Rectangle", "Rectangle");
            AddToolButton(toolStrip, "✏ Pen", "Pen");
            AddToolButton(toolStrip, "📝 Annotation", "Annotation");
            AddToolButton(toolStrip, "➖ Horizontal", "Horizontal");
            AddToolButton(toolStrip, "➕ Vertical", "Vertical");

            toolStrip.Items.Add(new ToolStripSeparator());

            ToolStripButton resetBtn = new ToolStripButton("❌ Reset");
            resetBtn.Click += (s, e) => _imageEditor.SetTool("");
            toolStrip.Items.Add(resetBtn);


            _saveHandler = new SaveOptionsHandler(btnSaveToWord, chkAppend, lblStatus, bottomPanel);
            _saveHandler.AttachSaveCloseButton(btnSaveClose);
            _saveHandler.SetImageProvider(() => _imageEditor.GetEditedImage());



            // Create menu holder in form
            MenuStrip menu = MainMenuBuilder.BuildMenu(this);
            Controls.Add(menu);
            MainMenuStrip = menu;
        }



        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            lblStatus.Text = "Capturing full screen...";
            WindowState = FormWindowState.Minimized;
            await Task.Delay(500);

            Rectangle bounds = GetPhysicalScreenBounds();
            Bitmap bmp = new Bitmap(bounds.Width, bounds.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
            }

            _currentCapture?.Dispose();
            _currentCapture = bmp;
            _imageEditor.SetImage(_currentCapture);
            _saveHandler.NotifyImageAvailable(true);

            //Clipboard.SetImage(_currentCapture);
            lblStatus.Text = "Full screen captured.";
            WindowState = FormWindowState.Normal;
        }

        private void BtnRegionCapture_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
            System.Threading.Thread.Sleep(300);

            using (RegionCaptureForm f = new RegionCaptureForm())
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    _currentCapture?.Dispose();
                    _currentCapture = f.CapturedBitmap;
                    _imageEditor.SetImage(_currentCapture);
                    _saveHandler.NotifyImageAvailable(true);
                    //Clipboard.SetImage(_currentCapture);
                    lblStatus.Text = "Region captured.";
                }
            }
            WindowState = FormWindowState.Normal;
        }

        private Rectangle GetPhysicalScreenBounds()
        {
            int left = int.MaxValue, top = int.MaxValue, right = int.MinValue, bottom = int.MinValue;
            foreach (Screen s in Screen.AllScreens)
            {
                left = Math.Min(left, s.Bounds.Left);
                top = Math.Min(top, s.Bounds.Top);
                right = Math.Max(right, s.Bounds.Right);
                bottom = Math.Max(bottom, s.Bounds.Bottom);
            }
            return Rectangle.FromLTRB(left, top, right, bottom);
        }

        private void BtnSetLocation_Click(object sender, EventArgs e)
        {
            SaveLocationManager.PromptAndSetFolder();
        }

            //helpder method for tool bar 
        private ToolStripButton AddToolButton(
    ToolStrip strip,
    string text,
    string toolName)
        {
            ToolStripButton btn = new ToolStripButton(text)
            {
                DisplayStyle = ToolStripItemDisplayStyle.Text,
                CheckOnClick = true
            };

            btn.Click += (s, e) =>
            {
                // Uncheck others
                foreach (ToolStripItem item in strip.Items)
                    if (item is ToolStripButton b) b.Checked = false;

                btn.Checked = true;
                _imageEditor.SetTool(toolName);
            };

            strip.Items.Add(btn);
            return btn;
        }


    }
}


using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        private ImageEditor _imageEditor;
        private SaveOptionsHandler _saveHandler;
        private Bitmap _currentCapture;

        public MainForm()
        {
            Text = "QA Evidence Capturer v1.6";
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

            lblStatus = new Label
            {
                Text = "Ready.",
                Location = new Point(20, 100),
                AutoSize = true,
                ForeColor = Color.DarkBlue
            };
            bottomPanel.Controls.Add(lblStatus);

            _saveHandler = new SaveOptionsHandler(btnSaveToWord, chkAppend, lblStatus, bottomPanel);
            _saveHandler.AttachSaveCloseButton(btnSaveClose);
            _saveHandler.SetImageProvider(() => _imageEditor.GetEditedImage());
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
    }
}

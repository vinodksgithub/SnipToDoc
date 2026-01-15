using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Marshal = System.Runtime.InteropServices.Marshal;
using Word = Microsoft.Office.Interop.Word;
using Label = System.Windows.Forms.Label;
using Font = System.Drawing.Font;
using App = System.Windows.Forms.Application;

namespace ScreenCaptureUtility
{
    public partial class MainForm : Form
    {
        // UI Controls
        private PictureBox pictureBoxPreview;
        private Button btnCapture;
        private Button btnRegionCapture;
        private Button btnSaveToWord;
        private Button btnBrowse;
        private CheckBox chkAppend;
        private Label lblStatus;
        private Panel bottomPanel;

        // Data members
        private Bitmap _currentCapture;
        private SaveOptionsHandler _saveHandler;

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        public MainForm()
        {
            SetProcessDPIAware();

            this.Text = "QA Evidence Capturer v1.4";
            this.Size = new Size(1000, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(700, 450);

            InitializeUnifiedUI();
        }

        private void InitializeUnifiedUI()
        {
            // Bottom panel
            bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 150,
                BackColor = Color.FromArgb(230, 230, 230),
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(bottomPanel);

            // Preview box
            pictureBoxPreview = new PictureBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.Fixed3D,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Gray
            };
            this.Controls.Add(pictureBoxPreview);

            // Capture screen
            btnCapture = new Button
            {
                Text = "📸 Capture Screen",
                Location = new Point(20, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnCapture.Click += BtnCapture_Click;

            // Stay on Top checkbox
            CheckBox chkStayOnTop = new CheckBox
            {
                Text = "Stay on Top",
                Location = new Point(640, 55),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };
            chkStayOnTop.CheckedChanged += (s, e) =>
            {
                this.TopMost = chkStayOnTop.Checked;
                lblStatus.Text = chkStayOnTop.Checked ? "Window will stay on top." : "Window normal.";
            };
            bottomPanel.Controls.Add(chkStayOnTop);

            bottomPanel.Controls.Add(btnCapture);

            // Region capture
            btnRegionCapture = new Button
            {
                Text = "📐 Region Capture",
                Location = new Point(180, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnRegionCapture.Click += BtnRegionCapture_Click;
            bottomPanel.Controls.Add(btnRegionCapture);

            // Save button
            btnSaveToWord = new Button
            {
                Text = "💾 Save",
                Location = new Point(340, 15),
                Size = new Size(150, 40),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Enabled = false
            };
            bottomPanel.Controls.Add(btnSaveToWord);

            // Append checkbox
            chkAppend = new CheckBox
            {
                Text = "Append Mode (Keep Word Open)",
                Location = new Point(640, 25),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };
            bottomPanel.Controls.Add(chkAppend);

            // Status label
            lblStatus = new Label
            {
                Text = "Ready.",
                Location = new Point(20, 100),
                AutoSize = true,
                ForeColor = Color.DarkBlue
            };
            bottomPanel.Controls.Add(lblStatus);

            bottomPanel.BringToFront();
            pictureBoxPreview.SendToBack();

            // 👉 Initialize SaveOptionsHandler AFTER controls exist
            _saveHandler = new SaveOptionsHandler(btnSaveToWord, chkAppend, lblStatus, bottomPanel);
        }

        // ---------- FULL SCREEN CAPTURE ----------
        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = "Capturing full screen...";
                this.WindowState = FormWindowState.Minimized;
                await Task.Delay(400);

                Rectangle bounds = Screen.PrimaryScreen.Bounds;
                Bitmap bmp = new Bitmap(bounds.Width, bounds.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                _currentCapture?.Dispose();
                _currentCapture = bmp;

                pictureBoxPreview.Image = _currentCapture;
                Clipboard.SetImage(_currentCapture);

                _saveHandler.SetCapture(_currentCapture);
                lblStatus.Text = "Full screen captured and copied to clipboard.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Capture error: " + ex.Message);
            }
            finally
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        // ---------- REGION CAPTURE ----------
        private void BtnRegionCapture_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = "Select region...";
                this.WindowState = FormWindowState.Minimized;
                System.Threading.Thread.Sleep(300);

                using (RegionCaptureForm captureForm = new RegionCaptureForm())
                {
                    if (captureForm.ShowDialog() == DialogResult.OK &&
                        captureForm.CapturedBitmap != null)
                    {
                        _currentCapture?.Dispose();
                        _currentCapture = captureForm.CapturedBitmap;

                        pictureBoxPreview.Image = _currentCapture;
                        Clipboard.SetImage(_currentCapture);

                        _saveHandler.SetCapture(_currentCapture);
                        lblStatus.Text = "Region captured and copied to clipboard.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Region capture failed: " + ex.Message);
            }
            finally
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
        }
    }

    static class MainProgram
    {
        [STAThread]
        static void Main()
        {
            App.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            App.EnableVisualStyles();
            App.SetCompatibleTextRenderingDefault(false);
            App.Run(new MainForm());
        }
    }
}
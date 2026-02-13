using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
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

        // Changed from Label → LinkLabel
        private LinkLabel lblStatus;

        private Panel bottomPanel;
        private Button btnSetLocation;

        private ImageEditor _imageEditor;
        private SaveOptionsHandler _saveHandler;
        private Bitmap _currentCapture;

        // Delay (in seconds) before capture starts
        private int captureDelaySeconds = 0; // default = 0s

        public ImageEditor ImageEditor => _imageEditor;

        private void MainForm_Load(object sender, EventArgs e)
        {
            captureDelaySeconds = 0;
        }

        public MainForm()
        {
            Text = "QA Evidence Capturer v1.62";
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
                Text = "💾 Save ",
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

            // ===== Status (LinkLabel) =====
            lblStatus = new LinkLabel
            {
                Text = "Ready.  Click here to navigate evidence path",
                Location = new Point(20, 100),
                AutoSize = true,
                LinkColor = Color.Blue,
                ActiveLinkColor = Color.Red,
                VisitedLinkColor = Color.Purple
            };

            // Add link ONLY ONCE (important)
            int linkStart = lblStatus.Text.IndexOf("Click");
            lblStatus.Links.Add(linkStart,
                "Click here to navigate evidence path".Length);

            lblStatus.LinkClicked += LblStatus_LinkClicked;

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

            _saveHandler = new SaveOptionsHandler(btnSaveToWord, chkAppend, (Label)(object)lblStatus, bottomPanel);
            _saveHandler.AttachSaveCloseButton(btnSaveClose);
            _saveHandler.SetImageProvider(() => _imageEditor.GetEditedImage());

            // Create menu holder in form
            MenuStrip menu = MainMenuBuilder.BuildMenu(this);

            // ===== Delay Menu =====
            ToolStripMenuItem delayMenu = new ToolStripMenuItem("Delay");
            delayMenu.DropDownItems.Add("0s", null, DelayMenu_Click);
            delayMenu.DropDownItems.Add("1s", null, DelayMenu_Click);
            delayMenu.DropDownItems.Add("3s", null, DelayMenu_Click);
            delayMenu.DropDownItems.Add("5s", null, DelayMenu_Click);
            menu.Items.Add(delayMenu);

            Controls.Add(menu);
            MainMenuStrip = menu;
        }

        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                btnCapture.Enabled = false;

                if (captureDelaySeconds > 0)
                {
                    lblStatus.Text = $"Capturing full screen in {captureDelaySeconds}s...";
                    await Task.Delay(captureDelaySeconds * 1000);
                }

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

                lblStatus.Text = "Full screen captured.";

                if (captureDelaySeconds > 0)
                {
                    ShowToast(" Capture completed");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Capture Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                WindowState = FormWindowState.Normal;
                btnCapture.Enabled = true;
            }
        }

        private async void BtnRegionCapture_Click(object sender, EventArgs e)
        {
            try
            {
                btnRegionCapture.Enabled = false;

                if (captureDelaySeconds > 0)
                {
                    lblStatus.Text = $"Region capture in {captureDelaySeconds}s...";
                    await Task.Delay(captureDelaySeconds * 1000);
                }

                WindowState = FormWindowState.Minimized;
                await Task.Delay(300);

                using (RegionCaptureForm f = new RegionCaptureForm())
                {
                    if (f.ShowDialog() == DialogResult.OK)
                    {
                        _currentCapture?.Dispose();
                        _currentCapture = f.CapturedBitmap;

                        _imageEditor.SetImage(_currentCapture);
                        _saveHandler.NotifyImageAvailable(true);

                        lblStatus.Text = "Region captured.";

                        if (captureDelaySeconds > 0)
                        {
                            ShowToast("✅ Capture completed");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Region Capture Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                WindowState = FormWindowState.Normal;
                btnRegionCapture.Enabled = true;
            }
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

            string text = "Ready.  Click here to navigate evidence path";

            lblStatus.SuspendLayout();

            try
            {
                lblStatus.Text = text;

                // Rebuild links safely
                lblStatus.Links.Clear();

                int linkStart = text.IndexOf("Click");
                if (linkStart >= 0)
                {
                    lblStatus.Links.Add(
                        linkStart,
                        "Click here to navigate evidence path".Length);
                }
            }
            catch (Exception ex)
            {
                // Optional: log/debug instead of silent failure
                System.Diagnostics.Debug.WriteLine(
                    $"LinkLabel update failed: {ex.Message}");
            }
            finally
            {
                // ALWAYS resume layout
                lblStatus.ResumeLayout();
            }
        }




        private void LblStatus_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string folder = SaveLocationManager.GetSaveFolder();

            if (System.IO.Directory.Exists(folder))
            {
                System.Diagnostics.Process.Start("explorer.exe", folder);
            }
        }


        // Helper method for tool bar
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
                foreach (ToolStripItem item in strip.Items)
                    if (item is ToolStripButton b) b.Checked = false;

                btn.Checked = true;
                _imageEditor.SetTool(toolName);
            };

            strip.Items.Add(btn);
            return btn;
        }

        // ===== Delay Menu Handler =====
        public void DelayMenu_Click(object sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem selectedItem)
                return;

            if (selectedItem.OwnerItem is not ToolStripMenuItem parentMenu)
                return;

            foreach (ToolStripMenuItem item in parentMenu.DropDownItems)
            {
                item.Checked = false;
            }

            selectedItem.Checked = true;

            captureDelaySeconds = int.Parse(
                selectedItem.Text.Replace("s", "")
            );
        }

        private async void ShowToast(string message)
        {
            Label toast = new Label
            {
                Text = message,
                AutoSize = true,
                BackColor = Color.FromArgb(180, 0, 0, 0),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Padding = new Padding(15),
                Visible = false
            };

            Controls.Add(toast);
            toast.BringToFront();

            toast.Left = ClientSize.Width - toast.Width - 20;
            toast.Top = ClientSize.Height - toast.Height - 20;

            toast.Visible = true;

            await Task.Delay(1000);

            for (int alpha = 180; alpha >= 0; alpha -= 30)
            {
                toast.BackColor = Color.FromArgb(alpha, 0, 0, 0);
                await Task.Delay(50);
            }

            Controls.Remove(toast);
            toast.Dispose();
        }
    }
}

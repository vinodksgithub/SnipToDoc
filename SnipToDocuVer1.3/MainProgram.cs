using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
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
        private Word.Application _wordApp = null;
        private Word.Document _activeDoc = null;
        private string _selectedWordPath = null;

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        public MainForm()
        {
            SetProcessDPIAware();

            this.Text = "QA Evidence Capturer v1.2";
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
            btnSaveToWord.Click += BtnSaveToWord_Click;
            bottomPanel.Controls.Add(btnSaveToWord);

            // Radio buttons for save options
            RadioButton radioWord = new RadioButton
            {
                Text = "Word",
                Location = new Point(340, 60),
                AutoSize = true,
                Checked = true // default option
            };
            bottomPanel.Controls.Add(radioWord);

            RadioButton radioImage = new RadioButton
            {
                Text = "Image",
                Location = new Point(400, 60),
                AutoSize = true
            };
            bottomPanel.Controls.Add(radioImage);

            // Browse
            btnBrowse = new Button
            {
                Text = "📂 Browse",
                Location = new Point(500, 15),
                Size = new Size(120, 40),
                Font = new Font("Segoe UI", 9)
            };
            btnBrowse.Click += BtnBrowse_Click;
            bottomPanel.Controls.Add(btnBrowse);

            // Append checkbox
            chkAppend = new CheckBox
            {
                Text = "Append Mode (Keep Word Open)",
                Location = new Point(640, 25),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };
            chkAppend.CheckedChanged += (s, e) =>
            {
                if (!chkAppend.Checked)
                {
                    CloseAndCleanupWord();
                    _selectedWordPath = null;
                }
                btnSaveToWord.Text = chkAppend.Checked ? "➕ Append to Doc" : "💾 Save to Word";
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

            // Hook save button to radio logic
            btnSaveToWord.Click += (s, e) =>
            {
                if (_currentCapture == null) return;

                if (radioWord.Checked)
                {
                    BtnSaveToWord_Click(s, e); // existing Word save logic
                }
                else if (radioImage.Checked)
                {
                    using (SaveFileDialog sfd = new SaveFileDialog())
                    {
                        sfd.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg";
                        sfd.Title = "Save Capture as Image";
                        sfd.FileName = $"Capture_{DateTime.Now:yyyyMMdd_HHmmss}.png";

                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            var format = sfd.FilterIndex == 2 ? ImageFormat.Jpeg : ImageFormat.Png;
                            _currentCapture.Save(sfd.FileName, format);
                            lblStatus.Text = $"Image saved: {Path.GetFileName(sfd.FileName)}";
                        }
                    }
                }
            };
        }

        // ---------- FULL SCREEN CAPTURE ----------
        // ---------- FULL SCREEN CAPTURE ----------
        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = "Capturing full screen...";
                this.WindowState = FormWindowState.Minimized;
                await Task.Delay(400);

                // Capture the bounds of the primary screen
                Rectangle bounds = Screen.PrimaryScreen.Bounds;
                Bitmap bmp = new Bitmap(bounds.Width, bounds.Height);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                // Dispose old capture if any
                _currentCapture?.Dispose();
                _currentCapture = bmp;

                // Show in preview
                pictureBoxPreview.Image = _currentCapture;

                // NEW: Copy to clipboard
                Clipboard.SetImage(_currentCapture);

                // Enable save button
                btnSaveToWord.Enabled = true;
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
                        // Dispose old capture if any
                        _currentCapture?.Dispose();
                        _currentCapture = captureForm.CapturedBitmap;

                        // Show in preview
                        pictureBoxPreview.Image = _currentCapture;

                        // NEW: Copy to clipboard
                        Clipboard.SetImage(_currentCapture);

                        // Enable save button
                        btnSaveToWord.Enabled = true;
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

        // ---------- BROWSE ----------
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Word Document (*.docx)|*.docx";
                sfd.Title = "Select Word File Location";
                sfd.FileName = $"Evidence_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    _selectedWordPath = sfd.FileName;
                    lblStatus.Text = "Save location selected.";
                }
            }
        }

        // ---------- SAVE TO WORD ----------
        private void BtnSaveToWord_Click(object sender, EventArgs e)
        {
            if (_currentCapture == null) return;

            string tempImg = Path.Combine(Path.GetTempPath(), "evidence.png");
            _currentCapture.Save(tempImg, ImageFormat.Png);

            try
            {
                if (_wordApp == null)
                {
                    _wordApp = new Word.Application { Visible = true };
                }

                if (_activeDoc == null)
                {
                    _activeDoc = _wordApp.Documents.Add();
                    _activeDoc.Content.InsertAfter("Test Execution Evidence Report\n");
                    _activeDoc.Content.InsertAfter($"Generated: {DateTime.Now}\n\n");
                }

                var range = _activeDoc.Content;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertAfter($"\nCaptured: {DateTime.Now}\n");

                var shape = _activeDoc.InlineShapes.AddPicture(
                    tempImg, Range: _activeDoc.Characters.Last);
                shape.Width = 450;

                if (!chkAppend.Checked)
                {
                    // Always generate a unique filename
                    string uniqueFileName = $"Evidence_{DateTime.Now:yyyyMMdd_HHmmss_fff}.docx";

                    // If user browsed a folder, use that folder; otherwise default to Desktop
                    string folderPath = string.IsNullOrEmpty(_selectedWordPath)
                        ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                        : Path.GetDirectoryName(_selectedWordPath);

                    string fullPath = Path.Combine(folderPath, uniqueFileName);

                    _activeDoc.SaveAs2(fullPath);
                    CloseAndCleanupWord();
                    lblStatus.Text = $"Saved successfully: {uniqueFileName}";
                }
                else
                {
                    lblStatus.Text = "Appended to active document.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Word error: " + ex.Message);
                CloseAndCleanupWord();
            }
            finally
            {
                if (File.Exists(tempImg))
                    File.Delete(tempImg);
            }
        }

        // ---------- CLEANUP ----------
        private void CloseAndCleanupWord()
        {
            try
            {
                if (_activeDoc != null)
                {
                    Marshal.ReleaseComObject(_activeDoc);
                    _activeDoc = null;
                }
                if (_wordApp != null)
                {
                    _wordApp.Quit();
                    Marshal.ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            }
            catch { }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            CloseAndCleanupWord();
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

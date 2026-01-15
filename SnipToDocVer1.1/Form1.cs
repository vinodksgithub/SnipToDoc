using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Word = Microsoft.Office.Interop.Word;

namespace ScreenCaptureUtility
{
    public partial class MainForm1 : Form
    {
        // UI Controls
        private PictureBox pictureBoxPreview;
        private Button btnCapture;
        private Button btnSaveToWord;
        private System.Windows.Forms.Label lblStatus;

        // Store the captured image in memory
        private Bitmap _currentCapture;

        // DPI Awareness Import
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        public MainForm1()
        {
            SetProcessDPIAware(); // Fix resolution for high DPI screens
            InitializeCustomUI(); // Build the UI
        }

        // ---------------------------------------------------------
        // 1. UI SETUP (Programmatic)
        // ---------------------------------------------------------
        private void InitializeCustomUI()
        {
            this.Text = "Test Evidence Capturer";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Preview Box
            pictureBoxPreview = new PictureBox();
            pictureBoxPreview.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxPreview.SizeMode = PictureBoxSizeMode.Zoom; // Important for preview!
            pictureBoxPreview.Location = new Point(12, 12);
            pictureBoxPreview.Size = new Size(760, 480);
            pictureBoxPreview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(pictureBoxPreview);

            // Capture Button
            btnCapture = new Button();
            btnCapture.Text = "Capture Screen";
            btnCapture.Location = new Point(12, 510);
            btnCapture.Size = new Size(120, 40);
            btnCapture.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCapture.Click += BtnCapture_Click;
            this.Controls.Add(btnCapture);

            // Save Button
            btnSaveToWord = new Button();
            btnSaveToWord.Text = "Save to Word";
            btnSaveToWord.Location = new Point(140, 510);
            btnSaveToWord.Size = new Size(120, 40);
            btnSaveToWord.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSaveToWord.Enabled = false; // Disabled until we have an image
            btnSaveToWord.Click += BtnSaveToWord_Click;
            this.Controls.Add(btnSaveToWord);

            // Status Label
            lblStatus = new System.Windows.Forms.Label();
            lblStatus.Text = "Ready.";
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(270, 523);
            lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.Controls.Add(lblStatus);
        }

        // ---------------------------------------------------------
        // 2. CAPTURE LOGIC
        // ---------------------------------------------------------
        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = "Capturing...";

                // 1. Hide the utility
                this.WindowState = FormWindowState.Minimized;

                // 2. Wait for the minimize animation to finish 
                // (Increase this to 700ms if you have a slow PC)
                await Task.Delay(700);

                // 3. Perform the capture
                // We use System.Windows.Forms.Screen to get the correct area
                Rectangle bounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;

                Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // The actual screen copying command
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                // 4. Update the UI
                if (_currentCapture != null) _currentCapture.Dispose(); // Clean up old image
                _currentCapture = bitmap;
                pictureBoxPreview.Image = _currentCapture;

                // 5. Success State
                btnSaveToWord.Enabled = true;
                lblStatus.Text = $"Captured {bounds.Width}x{bounds.Height}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Capture Failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 6. ALWAYS bring the window back
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }
        }

        private void InitializeComponent()
        {

        }

        // ---------------------------------------------------------
        // 3. WORD EXPORT LOGIC
        // ---------------------------------------------------------
        private void BtnSaveToWord_Click(object sender, EventArgs e)
        {
            if (_currentCapture == null) return;

            lblStatus.Text = "Generating Word Document...";
            System.Windows.Forms.Application.DoEvents(); // Force UI to update text immediately

            string tempFile = Path.Combine(Path.GetTempPath(), $"evidence_{Guid.NewGuid()}.png");

            // Generate a timestamped filename for the Word doc
            string docName = $"TestEvidence_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            string docPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), docName);

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                // 1. Save Bitmap to temp file (Word needs a file path usually)
                _currentCapture.Save(tempFile, ImageFormat.Png);

                // 2. Open Word
                wordApp = new Word.Application();
                wordApp.Visible = false; // Keep hidden while working

                // 3. Create Doc and Add Image
                doc = wordApp.Documents.Add();

                // Add a header text
                doc.Content.InsertAfter($"Test Execution Evidence - {DateTime.Now}\n");

                // Add the picture
                var range = doc.Content;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InlineShapes.AddPicture(tempFile, false, true);

                // 4. Save
                doc.SaveAs2(docPath);

                lblStatus.Text = $"Saved to Desktop: {docName}";
                MessageBox.Show($"Evidence saved successfully!\n\nPath: {docPath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Error saving document.";
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // 5. Cleanup Resources
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                // Cleanup temp image
                if (File.Exists(tempFile))
                    File.Delete(tempFile);
            }
        }
    }
}
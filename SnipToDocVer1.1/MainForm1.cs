using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ScreenCaptureUtility
{
    public partial class MainForm1 : Form
    {
        // UI Controls
        private PictureBox pictureBoxPreview;
        private Button btnCapture;
        private Button btnSaveToWord;
        private Label lblStatus;

        // Store the captured image in memory
        private Bitmap _currentCapture;

        // DLL Import to ensure high-resolution screens don't look blurry or cropped
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        public MainForm1()
        {
            SetProcessDPIAware();
            InitializeCustomUI();
        }

        private void InitializeCustomUI()
        {
            // Form Settings
            this.Text = "QA Test Evidence Utility";
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);

            // Preview Box
            pictureBoxPreview = new PictureBox();
            pictureBoxPreview.BorderStyle = BorderStyle.Fixed3D;
            pictureBoxPreview.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBoxPreview.Location = new Point(20, 20);
            pictureBoxPreview.Size = new Size(840, 550);
            pictureBoxPreview.BackColor = Color.White;
            pictureBoxPreview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(pictureBoxPreview);

            // Capture Button
            btnCapture = new Button();
            btnCapture.Text = "📸 Capture Screen";
            btnCapture.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnCapture.Location = new Point(20, 590);
            btnCapture.Size = new Size(160, 45);
            btnCapture.BackColor = Color.LightSkyBlue;
            btnCapture.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCapture.Click += BtnCapture_Click;
            this.Controls.Add(btnCapture);

            // Save Button
            btnSaveToWord = new Button();
            btnSaveToWord.Text = "💾 Save to Word";
            btnSaveToWord.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnSaveToWord.Location = new Point(190, 590);
            btnSaveToWord.Size = new Size(160, 45);
            btnSaveToWord.BackColor = Color.LightGreen;
            btnSaveToWord.Enabled = false;
            btnSaveToWord.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSaveToWord.Click += BtnSaveToWord_Click;
            this.Controls.Add(btnSaveToWord);

            // Status Label
            lblStatus = new Label();
            lblStatus.Text = "Ready to capture evidence.";
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Segoe UI", 9, FontStyle.Italic);
            lblStatus.Location = new Point(370, 605);
            lblStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            this.Controls.Add(lblStatus);
        }

        private async void BtnCapture_Click(object sender, EventArgs e)
        {
            try
            {
                lblStatus.Text = "Minimizing and capturing...";
                this.WindowState = FormWindowState.Minimized;

                // Wait for the window to vanish completely
                await Task.Delay(500);

                // Determine the primary screen size
                Rectangle bounds = Screen.PrimaryScreen.Bounds;

                // Create the bitmap
                Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);

                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Copy the screen onto the bitmap
                    g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                }

                // Update the memory and UI
                if (_currentCapture != null) _currentCapture.Dispose();
                _currentCapture = bitmap;
                pictureBoxPreview.Image = _currentCapture;
                pictureBoxPreview.Refresh(); // Force the box to show the image

                btnSaveToWord.Enabled = true;
                lblStatus.Text = $"Last capture: {DateTime.Now.ToLongTimeString()}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Restore window
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }
        }

        private void BtnSaveToWord_Click(object sender, EventArgs e)
        {
            if (_currentCapture == null) return;

            string tempPath = Path.Combine(Path.GetTempPath(), "temp_evidence.png");
            string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"Evidence_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                lblStatus.Text = "Generating Word file...";
                _currentCapture.Save(tempPath, ImageFormat.Png);

                wordApp = new Word.Application();
                doc = wordApp.Documents.Add();

                // Add Title
                var titleRange = doc.Range();
                titleRange.Text = $"Test Evidence Capture\nDate: {DateTime.Now}\n\n";
                titleRange.Font.Name = "Arial";
                titleRange.Font.Size = 16;
                titleRange.InsertParagraphAfter();

                // Add Screenshot
                var shape = doc.InlineShapes.AddPicture(tempPath);
                shape.Width = 450; // Scale to fit page

                doc.SaveAs2(desktopPath);
                lblStatus.Text = "Saved to Desktop!";
                MessageBox.Show($"File saved to: {desktopPath}", "Success");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Word Error: " + ex.Message);
            }
            finally
            {
                if (doc != null) { doc.Close(); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
                if (File.Exists(tempPath)) File.Delete(tempPath);
            }
        }
    }

    // THE ENTRY POINT CLASS
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
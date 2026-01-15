using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Marshal = System.Runtime.InteropServices.Marshal;
using Word = Microsoft.Office.Interop.Word;
using Label = System.Windows.Forms.Label;
using SnipToDocVer1._7.Properties;

namespace ScreenCaptureUtility
{
    public class SaveOptionsHandler
    {
        private RadioButton radioWord;
        private RadioButton radioImage;
        private Button btnSave;
        private Button btnSaveClose;
        private CheckBox chkAppend;
        private Label lblStatus;

        private Word.Application _wordApp = null;
        private Word.Document _activeDoc = null;

        private Func<Bitmap> _getEditedImage;

        public SaveOptionsHandler(Button saveButton, CheckBox appendCheckBox, Label statusLabel, Panel parentPanel)
        {
            btnSave = saveButton;
            chkAppend = appendCheckBox;
            lblStatus = statusLabel;

            radioWord = new RadioButton
            {
                Text = "Word",
                Location = new Point(saveButton.Left, saveButton.Bottom + 10),
                AutoSize = true,
                Checked = true
            };
            parentPanel.Controls.Add(radioWord);

            radioImage = new RadioButton
            {
                Text = "Image",
                Location = new Point(radioWord.Right + 20, saveButton.Bottom + 10),
                AutoSize = true
            };
            parentPanel.Controls.Add(radioImage);

            btnSave.Click += BtnSave_Click;

            chkAppend.CheckedChanged += (s, e) =>
            {
                if (chkAppend.Checked)
                {
                    btnSave.Text = "➕ Append";
                    btnSaveClose.Visible = true;

                    // Hide format selection in Append mode
                    radioWord.Visible = false;
                    radioImage.Visible = false;

                    // Always Word in append
                    radioWord.Checked = true;
                }
                else
                {
                    btnSave.Text = "💾 Save";
                    btnSaveClose.Visible = false;

                    // Restore options
                    radioWord.Visible = true;
                    radioImage.Visible = true;
                }
            };
        }

        public void AttachSaveCloseButton(Button saveCloseBtn)
        {
            btnSaveClose = saveCloseBtn;
            btnSaveClose.Click += BtnSaveClose_Click;
        }

        public void SetImageProvider(Func<Bitmap> imageProvider)
        {
            _getEditedImage = imageProvider;
            btnSave.Enabled = false;
        }

        public void NotifyImageAvailable(bool available)
        {
            btnSave.Enabled = available;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_getEditedImage == null) return;
            Bitmap currentCapture = _getEditedImage();
            if (currentCapture == null) return;

            if (chkAppend.Checked || radioWord.Checked)
                SaveToWord(currentCapture);
            else
                SaveToImage(currentCapture);
        }

        private void BtnSaveClose_Click(object sender, EventArgs e)
        {
            if (_activeDoc == null || _wordApp == null)
            {
                MessageBox.Show("No open Word document found.");
                return;
            }

            try
            {
                string folder = Settings.Default.SaveFolder;
                if (string.IsNullOrEmpty(folder))
                    folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                string fileName = $"Evidence_{DateTime.Now:yyyyMMdd_HHmmss_fff}.docx";
                string fullPath = Path.Combine(folder, fileName);

                _activeDoc.SaveAs2(fullPath);
                Clipboard.SetText(fullPath);
                lblStatus.Text = $"Saved & copied path: {fileName}";

                _activeDoc.Close(false);
                _wordApp.Quit();

                Marshal.ReleaseComObject(_activeDoc);
                Marshal.ReleaseComObject(_wordApp);

                _activeDoc = null;
                _wordApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save & Close failed: " + ex.Message);
            }
        }

        private void SaveToWord(Bitmap capture)
        {
            string tempImg = Path.Combine(Path.GetTempPath(), "evidence.png");
            capture.Save(tempImg, ImageFormat.Png);

            try
            {
                if (_wordApp == null)
                    _wordApp = new Word.Application { Visible = true };

                if (_activeDoc == null)
                {
                    _activeDoc = _wordApp.Documents.Add();
                    _activeDoc.Content.InsertAfter("Test Execution Evidence Report\n");
                    _activeDoc.Content.InsertAfter($"Generated: {DateTime.Now}\n\n");
                }

                var range = _activeDoc.Content;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertAfter($"\nCaptured: {DateTime.Now}\n");
                var shape = _activeDoc.InlineShapes.AddPicture(tempImg, Range: _activeDoc.Characters.Last);
                shape.Width = 450;

                if (!chkAppend.Checked)
                {
                    string uniqueFileName = $"Evidence_{DateTime.Now:yyyyMMdd_HHmmss_fff}.docx";
                    string folderPath = Settings.Default.SaveFolder;
                    if (string.IsNullOrEmpty(folderPath))
                        folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    string fullPath = Path.Combine(folderPath, uniqueFileName);

                    _activeDoc.SaveAs2(fullPath);
                    Clipboard.SetText(fullPath);
                    lblStatus.Text = $"Saved & copied path: {uniqueFileName}";

                    CloseAndCleanupWord();
                }
                else
                {
                    lblStatus.Text = "Appended. Click 'Save & Close' when finished.";
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

        private void SaveToImage(Bitmap capture)
        {
            string defaultFolder = Settings.Default.SaveFolder;
            if (string.IsNullOrEmpty(defaultFolder))
                defaultFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = defaultFolder;
                sfd.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg";
                sfd.Title = "Save Capture as Image";
                sfd.FileName = $"Capture_{DateTime.Now:yyyyMMdd_HHmmss}.png";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    ImageFormat format = sfd.FilterIndex == 2 ? ImageFormat.Jpeg : ImageFormat.Png;
                    capture.Save(sfd.FileName, format);
                    lblStatus.Text = $"Image saved: {Path.GetFileName(sfd.FileName)}";
                }
            }
        }

        private void CloseAndCleanupWord()
        {
            try
            {
                if (_activeDoc != null)
                {
                    if (!_activeDoc.Saved) _activeDoc.Save();
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
    }
}
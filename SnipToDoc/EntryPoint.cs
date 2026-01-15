using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ScreenCaptureToWord
{
    class Program
    {
        // Make process DPI aware so resolution is correct
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main(string[] args)
        {
            // Ensure DPI awareness
            SetProcessDPIAware();

            Console.WriteLine("Waiting 5 seconds before capturing...");
            Thread.Sleep(5000);

            // Capture screen
            Bitmap screenshot = CaptureScreen();

            // Save screenshot to temp file
            string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "screenshot.png");
            screenshot.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);

            // Save screenshot into Word document
            SaveImageToWord(tempFile);

            Console.WriteLine("Screenshot captured and saved to Word document.");
        }

        static Bitmap CaptureScreen()
        {
            // If you want only the primary screen:
            Rectangle bounds = Screen.PrimaryScreen.Bounds;

            // If you want ALL monitors, uncomment this:
            /*
            Rectangle bounds = Rectangle.Empty;
            foreach (var screen in Screen.AllScreens)
            {
                bounds = Rectangle.Union(bounds, screen.Bounds);
            }
            */

            Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height);

            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
            }

            return bitmap;
        }

        static void SaveImageToWord(string imagePath)
        {
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            Word.Document doc = wordApp.Documents.Add();
            doc.InlineShapes.AddPicture(imagePath);

            string docPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Screenshot.docx"
            );

            doc.SaveAs2(docPath);
            doc.Close();
            wordApp.Quit();
        }
    }
}
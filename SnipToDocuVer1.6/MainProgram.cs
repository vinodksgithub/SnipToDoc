using System;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using App = System.Windows.Forms.Application;

namespace ScreenCaptureUtility
{
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

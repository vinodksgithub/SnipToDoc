using System;
using System.Windows.Forms;
using App = System.Windows.Forms.Application;


/*
 * MainProgram.cs
 * ----------------
 * Entry point for the ScreenCaptureUtility application.
 *
 * Responsibilities:
 *  - Defines the Main method as the starting point of the program.
 *  - Configures application-wide settings before launching the main form.
 */
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

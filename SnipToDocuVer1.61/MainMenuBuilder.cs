using System;
using System.Windows.Forms;

namespace ScreenCaptureUtility
{
    public static class MainMenuBuilder
    {
        public static MenuStrip BuildMenu(MainForm form)
        {
            // Create menu strip
            MenuStrip menuStrip = new MenuStrip();

            // Top-level "Connect" menu
            ToolStripMenuItem connectMenu = new ToolStripMenuItem("Connect");

            // Sub-item: Connect ADO
            ToolStripMenuItem connectAdoItem = new ToolStripMenuItem("Connect ADO");
            connectAdoItem.Click += (s, e) => OnConnectAdo(form);
            connectMenu.DropDownItems.Add(connectAdoItem);

            // Sub-item: Upload Evidence To Test
            ToolStripMenuItem uploadEvidenceItem = new ToolStripMenuItem("Upload Evidence To Test");
            uploadEvidenceItem.Click += (s, e) => OnUploadEvidence(form);
            connectMenu.DropDownItems.Add(uploadEvidenceItem);

            // Add to menu strip
            menuStrip.Items.Add(connectMenu);

            return menuStrip;
        }

        private static void OnConnectAdo(MainForm form)
        {
            // Placeholder: open DB connection dialog
            MessageBox.Show("Connect ADO clicked. Implement DB connection logic here.",
                            "Connect ADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void OnUploadEvidence(MainForm form)
        {
            // Placeholder: upload evidence logic
            MessageBox.Show("Upload Evidence clicked. Implement upload logic here.",
                            "Upload Evidence", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
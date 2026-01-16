using System;
using System.Windows.Forms;

namespace ScreenCaptureUtility
{
    public static class MainMenuBuilder
    {
        public static MenuStrip BuildMenu(MainForm form)
        {
            MenuStrip menuStrip = new MenuStrip();

            // ===== Edit Menu =====
            ToolStripMenuItem editMenu = new ToolStripMenuItem("Edit");

            ToolStripMenuItem undoItem = new ToolStripMenuItem("Undo");
            undoItem.ShortcutKeys = Keys.Control | Keys.Z;
            undoItem.Click += (s, e) => form.ImageEditor?.Undo();

            ToolStripMenuItem redoItem = new ToolStripMenuItem("Redo");
            redoItem.ShortcutKeys = Keys.Control | Keys.Y;
            redoItem.Click += (s, e) => form.ImageEditor?.Redo();

            editMenu.DropDownItems.Add(undoItem);
            editMenu.DropDownItems.Add(redoItem);

            // ===== Existing Connect Menu =====
            ToolStripMenuItem connectMenu = new ToolStripMenuItem("Connect");

            ToolStripMenuItem connectAdoItem = new ToolStripMenuItem("Connect ADO");
            connectAdoItem.Click += (s, e) => OnConnectAdo(form);
            connectMenu.DropDownItems.Add(connectAdoItem);

            ToolStripMenuItem uploadEvidenceItem = new ToolStripMenuItem("Upload Evidence To Test");
            uploadEvidenceItem.Click += (s, e) => OnUploadEvidence(form);
            connectMenu.DropDownItems.Add(uploadEvidenceItem);

            // Add menus
            menuStrip.Items.Add(editMenu);
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
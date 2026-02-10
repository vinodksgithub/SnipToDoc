using System;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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

            ToolStripMenuItem delayMenu = new ToolStripMenuItem("DelayCapture");

            ToolStripMenuItem delay0s = new ToolStripMenuItem("0s") { Checked = true };
            ToolStripMenuItem delay3s = new ToolStripMenuItem("3s");
            ToolStripMenuItem delay5s = new ToolStripMenuItem("5s");
            ToolStripMenuItem delay9s = new ToolStripMenuItem("9s");

            delay0s.Click += form.DelayMenu_Click;
            delay3s.Click += form.DelayMenu_Click;
            delay5s.Click += form.DelayMenu_Click;
            delay9s.Click += form.DelayMenu_Click;

            delayMenu.DropDownItems.AddRange(new ToolStripItem[]
            {
            delay0s, delay3s, delay5s, delay9s
            });

            menuStrip.Items.Add(delayMenu);

            // Add menus
            menuStrip.Items.Add(editMenu);
            menuStrip.Items.Add(connectMenu);

            return menuStrip;
        }


        private static void OnConnectAdo(MainForm form)
        {
            // Placeholder: open DB connection dialog
            MessageBox.Show("Connect ADO clicked.",
                            "Connect ADO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void OnUploadEvidence(MainForm form)
        {
            // Placeholder: upload evidence logic
            MessageBox.Show("Upload Evidence clicked. ",
                            "Upload Evidence", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
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

            // ===== Connect Menu =====
            ToolStripMenuItem connectMenu = new ToolStripMenuItem("Connect");

            ToolStripMenuItem connectAdoItem = new ToolStripMenuItem("Connect ADO");
            connectAdoItem.Click += (s, e) => OnConnectAdo(form);
            connectMenu.DropDownItems.Add(connectAdoItem);

            ToolStripMenuItem uploadEvidenceItem = new ToolStripMenuItem("Upload Evidence To Test");
            uploadEvidenceItem.Click += (s, e) => OnUploadEvidence(form);
            connectMenu.DropDownItems.Add(uploadEvidenceItem);

            // ===== Word Addons Menu =====
            ToolStripMenuItem wordAddonsMenu = new ToolStripMenuItem("Word Addons");

            ToolStripMenuItem insertTextItem = new ToolStripMenuItem("Insert text")
            {
                CheckOnClick = true // toggle on/off
            };
            wordAddonsMenu.DropDownItems.Add(insertTextItem);

            wordAddonsMenu.DropDownItems.Add("Heading 1", null, (s, e) => SaveOptionsHandler.SetHeadingLevel(1));
            wordAddonsMenu.DropDownItems.Add("Heading 2", null, (s, e) => SaveOptionsHandler.SetHeadingLevel(2));
            wordAddonsMenu.DropDownItems.Add("Heading 3", null, (s, e) => SaveOptionsHandler.SetHeadingLevel(3));

            // Attach InsertText toggle to SaveOptionsHandler
            SaveOptionsHandler.SetInsertTextProvider(() => insertTextItem.Checked);

            // ===== Add Menus to Strip =====
            menuStrip.Items.Add(editMenu);
            menuStrip.Items.Add(connectMenu);
            menuStrip.Items.Add(wordAddonsMenu);

            return menuStrip;
        }

        // Placeholder methods for Connect actions
        private static void OnConnectAdo(MainForm form)
        {
            MessageBox.Show("Connect ADO clicked.");
        }

        private static void OnUploadEvidence(MainForm form)
        {
            MessageBox.Show("Upload Evidence clicked.");
        }
    }
}

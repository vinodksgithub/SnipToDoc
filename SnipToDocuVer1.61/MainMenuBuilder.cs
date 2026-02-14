using System;
using System.Windows.Forms;

namespace ScreenCaptureUtility
{
    public static class MainMenuBuilder
    {
        private static ToolStripMenuItem heading1Item;
        private static ToolStripMenuItem heading2Item;
        private static ToolStripMenuItem heading3Item;

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

            // Add separator
            wordAddonsMenu.DropDownItems.Add(new ToolStripSeparator());

            // Create Heading menu items as radio buttons
            heading1Item = new ToolStripMenuItem("Heading 1")
            {
                CheckOnClick = true
            };
            heading1Item.Click += (s, e) => OnHeadingSelected(1, heading1Item, heading2Item, heading3Item);

            heading2Item = new ToolStripMenuItem("Heading 2")
            {
                CheckOnClick = true
            };
            heading2Item.Click += (s, e) => OnHeadingSelected(2, heading1Item, heading2Item, heading3Item);

            heading3Item = new ToolStripMenuItem("Heading 3")
            {
                CheckOnClick = true
            };
            heading3Item.Click += (s, e) => OnHeadingSelected(3, heading1Item, heading2Item, heading3Item);

            wordAddonsMenu.DropDownItems.Add(heading1Item);
            wordAddonsMenu.DropDownItems.Add(heading2Item);
            wordAddonsMenu.DropDownItems.Add(heading3Item);

            // Set Heading 3 as default
            heading3Item.Checked = true;
            SaveOptionsHandler.SetHeadingLevel(3);

            // Attach InsertText toggle to SaveOptionsHandler
            SaveOptionsHandler.SetInsertTextProvider(() => insertTextItem.Checked);

            // ===== Add Menus to Strip =====
            menuStrip.Items.Add(editMenu);
            menuStrip.Items.Add(connectMenu);
            menuStrip.Items.Add(wordAddonsMenu);

            return menuStrip;
        }

        private static void OnHeadingSelected(int level, params ToolStripMenuItem[] allHeadingItems)
        {
            // Uncheck all heading items first
            foreach (var item in allHeadingItems)
            {
                item.Checked = false;
            }

            // Check only the selected one
            allHeadingItems[level - 1].Checked = true;

            // Set the heading level in SaveOptionsHandler
            SaveOptionsHandler.SetHeadingLevel(level);
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

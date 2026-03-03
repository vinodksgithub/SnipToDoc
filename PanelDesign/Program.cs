using System;
using System.Drawing;
using System.Reflection.Emit;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Font = System.Drawing.Font;
using Label = System.Windows.Forms.Label;
using Application = System.Windows.Forms.Application;

using System;
using System.Drawing;
using System.Windows.Forms;

namespace LayoutSample
{
    public class MainForm : Form
    {
        private ToolStripStatusLabel statusLabel;

        public MainForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            // =======================
            // FORM SETTINGS
            // =======================
            this.Text = "Layout Sample";
            this.Size = new Size(1600, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.Gainsboro;

            // =======================
            // MENU STRIP
            // =======================
            MenuStrip menuStrip = new MenuStrip();

            ToolStripMenuItem editMenu = new ToolStripMenuItem("Edit");
            editMenu.DropDownItems.Add("Copy");
            editMenu.DropDownItems.Add("Reset");

            ToolStripMenuItem captureMenu = new ToolStripMenuItem("Capture");
            captureMenu.DropDownItems.Add("Capture Full Screen");
            captureMenu.DropDownItems.Add("Capture Region");

            ToolStripMenuItem zoomMenu = new ToolStripMenuItem("Zoom");
            zoomMenu.DropDownItems.Add("Zoom In");
            zoomMenu.DropDownItems.Add("Zoom Out");

            ToolStripMenuItem helpMenu = new ToolStripMenuItem("Help");
            helpMenu.DropDownItems.Add("User Guide");
            helpMenu.DropDownItems.Add("About");

            menuStrip.Items.Add(editMenu);
            menuStrip.Items.Add(captureMenu);
            menuStrip.Items.Add(zoomMenu);
            menuStrip.Items.Add(helpMenu);

            this.MainMenuStrip = menuStrip;
            Controls.Add(menuStrip);

            // =======================
            // PICTURE AREA
            // =======================
            Panel picturePanel = new Panel
            {
                Location = new Point(50, 80),
                Size = new Size(1450, 520),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.WhiteSmoke
            };

            Label picText = new Label
            {
                Text = "Picture box",
                AutoSize = true,
                Location = new Point(160, 160),
                ForeColor = Color.DarkOrange,
                Font = new Font("Segoe UI", 14)
            };

            picturePanel.Controls.Add(picText);
            Controls.Add(picturePanel);

            // =======================
            // LEFT BUTTON GROUP
            // =======================
            Panel leftGroup = new Panel
            {
                Location = new Point(50, 630),
                Size = new Size(380, 80),
                BorderStyle = BorderStyle.FixedSingle
            };

            leftGroup.Controls.Add(CreateButton("Button 1", 20, 20, 120, 32));
            leftGroup.Controls.Add(CreateButton("Button 2", 200, 20, 120, 32));

            Controls.Add(leftGroup);

            // =======================
            // MIDDLE GROUP
            // =======================
            Panel middleGroup = new Panel
            {
                Location = new Point(470, 625),
                Size = new Size(140, 90),
                BorderStyle = BorderStyle.FixedSingle
            };

            Button btn3 = CreateButton("Button 3", 15, 12, 105, 30);

            RadioButton radio1 = new RadioButton
            {
                Location = new Point(25, 55),
                Size = new Size(20, 20)
            };

            RadioButton radio2 = new RadioButton
            {
                Location = new Point(85, 55),
                Size = new Size(20, 20)
            };

            middleGroup.Controls.Add(btn3);
            middleGroup.Controls.Add(radio1);
            middleGroup.Controls.Add(radio2);

            Controls.Add(middleGroup);

            // =======================
            // RIGHT SIDE LABELS + CHECKBOXES
            // =======================
            Label lbl4 = new Label
            {
                Text = "Option 1",
                Location = new Point(1200, 620),
                AutoSize = true,
                Font = new Font("Segoe UI", 11)
            };

            CheckBox chk4 = new CheckBox
            {
                Location = new Point(1320, 620),
                Size = new Size(20, 20)
            };

            Label lbl5 = new Label
            {
                Text = "Option 2",
                Location = new Point(1200, 660),
                AutoSize = true,
                Font = new Font("Segoe UI", 11)
            };

            CheckBox chk5 = new CheckBox
            {
                Location = new Point(1320, 660),
                Size = new Size(20, 20)
            };

            Controls.Add(lbl4);
            Controls.Add(chk4);
            Controls.Add(lbl5);
            Controls.Add(chk5);

            // =======================
            // CAPTURE PATH BUTTON
            // =======================
            Controls.Add(CreateButton("Capture Path", 1210, 760, 130, 38));

            // =======================
            // STATUS BAR
            // =======================
            StatusStrip statusStrip = new StatusStrip
            {
                Dock = DockStyle.Bottom
            };

            statusLabel = new ToolStripStatusLabel
            {
                Text = "Ready",
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft
            };

            statusStrip.Items.Add(statusLabel);
            Controls.Add(statusStrip);
        }

        // =======================
        // HELPER METHOD
        // =======================
        private Button CreateButton(string text, int x, int y, int w, int h)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                Font = new Font("Segoe UI", 9)
            };
        }

        // =======================
        // PROGRAM ENTRY
        // =======================
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}

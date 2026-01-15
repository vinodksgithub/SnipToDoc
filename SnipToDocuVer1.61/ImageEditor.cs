using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;

namespace ScreenCaptureUtility
{
    public class ImageEditor
    {
        private PictureBox _pictureBox;

        private Bitmap _baseImage;       // original screenshot
        private Bitmap _overlayImage;    // drawings only

        private bool _isDrawing;
        private Point _startImgPoint;
        private Point _lastImgPoint;

        private enum DrawMode { None, Rectangle, Pen, Text }
        private DrawMode _mode = DrawMode.None;

        private Rectangle _currentImgRect;
        private GraphicsPath _penPath;
        private ContextMenuStrip _menu;

        public ImageEditor(PictureBox pictureBox)
        {
            _pictureBox = pictureBox;
            _pictureBox.MouseDown += PictureBox_MouseDown;
            _pictureBox.MouseMove += PictureBox_MouseMove;
            _pictureBox.MouseUp += PictureBox_MouseUp;
            _pictureBox.Paint += PictureBox_Paint;
            BuildContextMenu();
        }

        private void BuildContextMenu()
        {
            _menu = new ContextMenuStrip();
            _menu.Items.Add("Rectangle", null, (s, e) => _mode = DrawMode.Rectangle);
            _menu.Items.Add("Pen", null, (s, e) => _mode = DrawMode.Pen);
            _menu.Items.Add("Annotation", null, (s, e) => _mode = DrawMode.Text);
            _menu.Items.Add(new ToolStripSeparator());

            //New Reset option
            _menu.Items.Add("Reset", null, (s, e) =>
            {
                _mode = DrawMode.None;
                _currentImgRect = Rectangle.Empty;
                _penPath = null;
                _isDrawing = false;
                _pictureBox.Invalidate();   // refresh UI
            });

            _menu.Items.Add("Erase All", null, (s, e) => EraseAll());
            _pictureBox.ContextMenuStrip = _menu;
        }

        public void SetImage(Bitmap bmp)
        {
            _baseImage = new Bitmap(bmp);
            _overlayImage = new Bitmap(bmp.Width, bmp.Height);
            RefreshComposite();
        }

        private void PictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (_baseImage == null) return;

            if (e.Button == MouseButtons.Right)
            {
                _menu.Show(_pictureBox, e.Location);
                return;
            }

            Point imgPoint = TranslateZoomMousePosition(e.Location);

            if (_mode == DrawMode.Text)
            {
                string text = Prompt.ShowDialog("Enter annotation text", "Add Annotation");
                if (!string.IsNullOrEmpty(text))
                {
                    using (Graphics g = Graphics.FromImage(_overlayImage))
                    using (Font f = new Font("Segoe UI", 14, FontStyle.Bold))
                    using (Brush b = new SolidBrush(Color.Red))
                    {
                        // Define a rectangle for text layout
                        RectangleF layoutRect = new RectangleF(
                            imgPoint.X, imgPoint.Y,
                            _baseImage.Width - imgPoint.X - 10,
                            _baseImage.Height - imgPoint.Y - 10);

                        // Configure wrapping and line breaks
                        StringFormat format = new StringFormat()
                        {
                            Alignment = StringAlignment.Near,
                            LineAlignment = StringAlignment.Near,
                            Trimming = StringTrimming.Word
                        };

                        g.DrawString(text, f, b, layoutRect, format);
                    }
                    RefreshComposite();
                }
                return;
            }

            _isDrawing = true;
            _startImgPoint = imgPoint;
            _lastImgPoint = imgPoint;

            if (_mode == DrawMode.Pen)
            {
                _penPath = new GraphicsPath();
                _penPath.AddLine(_startImgPoint, _startImgPoint);
            }
        }

        private void PictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDrawing) return;
            Point current = TranslateZoomMousePosition(e.Location);

            if (_mode == DrawMode.Pen)
            {
                _penPath.AddLine(_lastImgPoint, current);
                _lastImgPoint = current;
                _pictureBox.Invalidate();
            }
            else if (_mode == DrawMode.Rectangle)
            {
                _currentImgRect = GetRect(_startImgPoint, current);
                _pictureBox.Invalidate();
            }
        }

        private void PictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (!_isDrawing) return;

            if (_mode == DrawMode.Rectangle)
            {
                using (Graphics g = Graphics.FromImage(_overlayImage))
                using (Pen pen = new Pen(Color.Red, 2))
                {
                    g.DrawRectangle(pen, _currentImgRect);
                }
            }

            if (_mode == DrawMode.Pen && _penPath != null)
            {
                using (Graphics g = Graphics.FromImage(_overlayImage))
                using (Pen pen = new Pen(Color.Red, 2))
                {
                    g.DrawPath(pen, _penPath);
                }
                _penPath.Dispose();
                _penPath = null;
            }

            _isDrawing = false;
            _currentImgRect = Rectangle.Empty;
            RefreshComposite();
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            if (_isDrawing && _mode == DrawMode.Rectangle && _currentImgRect.Width > 0)
            {
                Rectangle drawRect = TranslateImageRectToControl(_currentImgRect);
                using (Pen pen = new Pen(Color.Red, 2))
                    e.Graphics.DrawRectangle(pen, drawRect);
            }

            if (_isDrawing && _mode == DrawMode.Pen && _penPath != null)
            {
                using (Pen pen = new Pen(Color.Red, 2))
                    e.Graphics.DrawPath(pen, TranslatePathToControl(_penPath));
            }
        }

        private void EraseAll()
        {
            _overlayImage = new Bitmap(_baseImage.Width, _baseImage.Height);
            RefreshComposite();
        }

        private void RefreshComposite()
        {
            Bitmap composite = new Bitmap(_baseImage.Width, _baseImage.Height);
            using (Graphics g = Graphics.FromImage(composite))
            {
                g.DrawImage(_baseImage, 0, 0);
                g.DrawImage(_overlayImage, 0, 0);
            }

            if (_pictureBox.Image != null)
                _pictureBox.Image.Dispose();

            _pictureBox.Image = composite;

            Clipboard.SetImage(new Bitmap(composite));
        }

        public Bitmap GetEditedImage() => new Bitmap(_pictureBox.Image);

        #region Coordinate helpers
        private Rectangle GetRect(Point p1, Point p2)
        {
            return new Rectangle(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y),
                Math.Abs(p1.X - p2.X), Math.Abs(p1.Y - p2.Y));
        }

        private Point TranslateZoomMousePosition(Point p)
        {
            float imageAspect = (float)_baseImage.Width / _baseImage.Height;
            float boxAspect = (float)_pictureBox.Width / _pictureBox.Height;

            int drawWidth, drawHeight, offsetX = 0, offsetY = 0;

            if (imageAspect > boxAspect)
            {
                drawWidth = _pictureBox.Width;
                drawHeight = (int)(_pictureBox.Width / imageAspect);
                offsetY = (_pictureBox.Height - drawHeight) / 2;
            }
            else
            {
                drawHeight = _pictureBox.Height;
                drawWidth = (int)(_pictureBox.Height * imageAspect);
                offsetX = (_pictureBox.Width - drawWidth) / 2;
            }

            float scaleX = (float)_baseImage.Width / drawWidth;
            float scaleY = (float)_baseImage.Height / drawHeight;

            return new Point((int)((p.X - offsetX) * scaleX), (int)((p.Y - offsetY) * scaleY));
        }

        private Rectangle TranslateImageRectToControl(Rectangle r)
        {
            float imageAspect = (float)_baseImage.Width / _baseImage.Height;
            float boxAspect = (float)_pictureBox.Width / _pictureBox.Height;

            int drawWidth, drawHeight, offsetX = 0, offsetY = 0;

            if (imageAspect > boxAspect)
            {
                drawWidth = _pictureBox.Width;
                drawHeight = (int)(_pictureBox.Width / imageAspect);
                offsetY = (_pictureBox.Height - drawHeight) / 2;
            }
            else
            {
                drawHeight = _pictureBox.Height;
                drawWidth = (int)(_pictureBox.Height * imageAspect);
                offsetX = (_pictureBox.Width - drawWidth) / 2;
            }

            float scaleX = (float)drawWidth / _baseImage.Width;
            float scaleY = (float)drawHeight / _baseImage.Height;

            return new Rectangle((int)(r.X * scaleX + offsetX), (int)(r.Y * scaleY + offsetY),
                                 (int)(r.Width * scaleX), (int)(r.Height * scaleY));
        }

        private GraphicsPath TranslatePathToControl(GraphicsPath imgPath)
        {
            GraphicsPath controlPath = new GraphicsPath();
            PointF[] pts = imgPath.PathPoints;

            for (int i = 1; i < pts.Length; i++)
            {
                controlPath.AddLine(
                    TranslateImageRectToControl(new Rectangle((int)pts[i - 1].X, (int)pts[i - 1].Y, 1, 1)).Location,
                    TranslateImageRectToControl(new Rectangle((int)pts[i].X, (int)pts[i].Y, 1, 1)).Location
                );
            }
            return controlPath;
        }
        #endregion
    }

    public static class Prompt
    {
        public static string ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 300,
                Text = caption,
                StartPosition = FormStartPosition.CenterParent
            };

            Label lbl = new Label()
            {
                Left = 20,
                Top = 20,
                Text = text,
                Width = 440
            };

            TextBox box = new TextBox()
            {
                Left = 20,
                Top = 50,
                Width = 440,
                Height = 150,
                Multiline = true,              //allow multiple lines
                ScrollBars = ScrollBars.Vertical, // add scrollbar
                AcceptsReturn = true           // allow Enter key for new lines
            };

            Button ok = new Button()
            {
                Text = "OK",
                Left = 380,
                Width = 80,
                Top = 220,
                DialogResult = DialogResult.OK
            };

            prompt.Controls.Add(lbl);
            prompt.Controls.Add(box);
            prompt.Controls.Add(ok);
            prompt.AcceptButton = ok;

            return prompt.ShowDialog() == DialogResult.OK ? box.Text : "";
        }
    }
}
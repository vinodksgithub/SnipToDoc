using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Label = System.Windows.Forms.Label;

namespace ScreenCaptureUtility
{
    public class ImageEditor
    {
        private PictureBox _pictureBox;

        private Bitmap _baseImage;
        private Bitmap _overlayImage;

        private bool _isDrawing;
        private Point _startImgPoint;
        private Point _lastImgPoint;
        private Point _currentLineEnd;

        private enum DrawMode
        {
            None,
            Rectangle,
            Pen,
            Text,
            HorizontalLine,
            VerticalLine
        }

        private DrawMode _mode = DrawMode.None;

        private Rectangle _currentImgRect;
        private GraphicsPath _penPath;

        private Stack<Bitmap> _undoStack = new Stack<Bitmap>();
        private Stack<Bitmap> _redoStack = new Stack<Bitmap>();

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

        #region Context Menu
        private void BuildContextMenu()
        {
            _menu = new ContextMenuStrip();

            _menu.Items.Add("Rectangle", null, (s, e) => _mode = DrawMode.Rectangle);
            _menu.Items.Add("Pen", null, (s, e) => _mode = DrawMode.Pen);
            _menu.Items.Add("Annotation", null, (s, e) => _mode = DrawMode.Text);
            _menu.Items.Add("Horizontal Line", null, (s, e) => _mode = DrawMode.HorizontalLine);
            _menu.Items.Add("Vertical Line", null, (s, e) => _mode = DrawMode.VerticalLine);

            _menu.Items.Add(new ToolStripSeparator());

            _menu.Items.Add("Reset Tool", null, (s, e) =>
            {
                _mode = DrawMode.None;
                _isDrawing = false;
            });

            _menu.Items.Add("Erase All", null, (s, e) => EraseAll());

            _pictureBox.ContextMenuStrip = _menu;
        }
        #endregion

        #region Public API (Used by Menu Bar)
        public void Undo()
        {
            if (_undoStack.Count == 0) return;

            _redoStack.Push(new Bitmap(_overlayImage));
            _overlayImage.Dispose();
            _overlayImage = _undoStack.Pop();

            RefreshComposite();
        }

        public void Redo()
        {
            if (_redoStack.Count == 0) return;

            _undoStack.Push(new Bitmap(_overlayImage));
            _overlayImage.Dispose();
            _overlayImage = _redoStack.Pop();

            RefreshComposite();
        }

        public Bitmap GetEditedImage()
        {
            return new Bitmap(_pictureBox.Image);
        }
        #endregion

        #region Image Setup
        public void SetImage(Bitmap bmp)
        {
            _baseImage?.Dispose();
            _overlayImage?.Dispose();

            _baseImage = new Bitmap(bmp);
            _overlayImage = new Bitmap(bmp.Width, bmp.Height);

            _undoStack.Clear();
            _redoStack.Clear();

            RefreshComposite();
        }
        #endregion

        #region Mouse Events
        private void PictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (_baseImage == null) return;

            if (e.Button == MouseButtons.Right) return;

            Point imgPoint = TranslateZoomMousePosition(e.Location);

            if (_mode == DrawMode.Text)
            {
                string text = Prompt.ShowDialog("Enter annotation text", "Add Annotation");
                if (!string.IsNullOrEmpty(text))
                {
                    SaveStateForUndo();

                    using (Graphics g = Graphics.FromImage(_overlayImage))
                    using (Font f = new Font("Segoe UI", 14, FontStyle.Bold))
                    using (Brush b = new SolidBrush(Color.Red))
                    {
                        g.DrawString(text, f, b, imgPoint);
                    }

                    RefreshComposite();
                }
                return;
            }

            _isDrawing = true;
            _startImgPoint = imgPoint;
            _lastImgPoint = imgPoint;
            _currentLineEnd = imgPoint;

            if (_mode == DrawMode.Pen)
            {
                _penPath = new GraphicsPath();
                _penPath.AddLine(imgPoint, imgPoint);
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
            }
            else if (_mode == DrawMode.Rectangle)
            {
                _currentImgRect = GetRect(_startImgPoint, current);
            }
            else if (_mode == DrawMode.HorizontalLine)
            {
                _currentLineEnd = new Point(current.X, _startImgPoint.Y);
            }
            else if (_mode == DrawMode.VerticalLine)
            {
                _currentLineEnd = new Point(_startImgPoint.X, current.Y);
            }

            _pictureBox.Invalidate();
        }

        private void PictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            if (!_isDrawing) return;

            SaveStateForUndo();

            using (Graphics g = Graphics.FromImage(_overlayImage))
            using (Pen pen = new Pen(Color.Red, 2))
            {
                if (_mode == DrawMode.Rectangle)
                    g.DrawRectangle(pen, _currentImgRect);

                else if (_mode == DrawMode.Pen && _penPath != null)
                    g.DrawPath(pen, _penPath);

                else if (_mode == DrawMode.HorizontalLine || _mode == DrawMode.VerticalLine)
                    g.DrawLine(pen, _startImgPoint, _currentLineEnd);
            }

            _penPath?.Dispose();
            _penPath = null;
            _currentImgRect = Rectangle.Empty;
            _isDrawing = false;

            RefreshComposite();
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            if (!_isDrawing) return;

            using (Pen pen = new Pen(Color.Red, 2))
            {
                if (_mode == DrawMode.Rectangle && _currentImgRect.Width > 0)
                    e.Graphics.DrawRectangle(pen, TranslateImageRectToControl(_currentImgRect));

                else if (_mode == DrawMode.Pen && _penPath != null)
                    e.Graphics.DrawPath(pen, TranslatePathToControl(_penPath));

                else if (_mode == DrawMode.HorizontalLine || _mode == DrawMode.VerticalLine)
                {
                    Point start = TranslateImagePointToControl(_startImgPoint);
                    Point end = TranslateImagePointToControl(_currentLineEnd);
                    e.Graphics.DrawLine(pen, start, end);
                }
            }
        }
        #endregion

        #region Undo Helpers
        private void SaveStateForUndo()
        {
            _undoStack.Push(new Bitmap(_overlayImage));
            _redoStack.Clear();
        }

        private void EraseAll()
        {
            SaveStateForUndo();
            _overlayImage = new Bitmap(_baseImage.Width, _baseImage.Height);
            RefreshComposite();
        }
        #endregion

        #region Rendering
        private void RefreshComposite()
        {
            Bitmap composite = new Bitmap(_baseImage.Width, _baseImage.Height);

            using (Graphics g = Graphics.FromImage(composite))
            {
                g.DrawImage(_baseImage, 0, 0);
                g.DrawImage(_overlayImage, 0, 0);
            }

            _pictureBox.Image?.Dispose();
            _pictureBox.Image = composite;

            Clipboard.SetImage(new Bitmap(composite));
        }
        #endregion

        #region Coordinate Helpers
        private Rectangle GetRect(Point p1, Point p2)
        {
            return new Rectangle(
                Math.Min(p1.X, p2.X),
                Math.Min(p1.Y, p2.Y),
                Math.Abs(p1.X - p2.X),
                Math.Abs(p1.Y - p2.Y));
        }

        private Point TranslateImagePointToControl(Point p)
        {
            Rectangle r = TranslateImageRectToControl(new Rectangle(p, new Size(1, 1)));
            return r.Location;
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

            return new Point(
                (int)((p.X - offsetX) * scaleX),
                (int)((p.Y - offsetY) * scaleY));
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

            return new Rectangle(
                (int)(r.X * scaleX + offsetX),
                (int)(r.Y * scaleY + offsetY),
                (int)(r.Width * scaleX),
                (int)(r.Height * scaleY));
        }

        private GraphicsPath TranslatePathToControl(GraphicsPath imgPath)
        {
            GraphicsPath controlPath = new GraphicsPath();
            PointF[] pts = imgPath.PathPoints;

            for (int i = 1; i < pts.Length; i++)
            {
                controlPath.AddLine(
                    TranslateImagePointToControl(Point.Round(pts[i - 1])),
                    TranslateImagePointToControl(Point.Round(pts[i])));
            }
            return controlPath;
        }
        #endregion
    }

    #region Prompt Dialog
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
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
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
    #endregion
}

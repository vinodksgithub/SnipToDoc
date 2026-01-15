using System;
using System.Drawing;
using System.Windows.Forms;

namespace ScreenCaptureUtility
{
    public class RegionCaptureForm : Form
    {
        private Point _startPoint;
        private Point _endPoint;
        private Rectangle _selection;
        private bool _isSelecting = false;

        public Bitmap CapturedBitmap { get; private set; }

        public RegionCaptureForm()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = Color.Black;
            this.Opacity = 0.25;
            this.TopMost = true;
            this.Cursor = Cursors.Cross;
            this.DoubleBuffered = true;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            _isSelecting = true;
            _startPoint = e.Location;
            _endPoint = e.Location;
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (_isSelecting)
            {
                _endPoint = e.Location;
                Invalidate();
            }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _isSelecting = false;

            _selection = GetRectangle(_startPoint, _endPoint);

            if (_selection.Width > 0 && _selection.Height > 0)
            {
                CapturedBitmap = new Bitmap(_selection.Width, _selection.Height);
                using (Graphics g = Graphics.FromImage(CapturedBitmap))
                {
                    g.CopyFromScreen(_selection.Location, Point.Empty, _selection.Size);
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (_isSelecting)
            {
                Rectangle rect = GetRectangle(_startPoint, _endPoint);
                using (Pen pen = new Pen(Color.Red, 6))
                {
                    e.Graphics.DrawRectangle(pen, rect);
                }
            }
        }

        private Rectangle GetRectangle(Point p1, Point p2)
        {
            return new Rectangle(
                Math.Min(p1.X, p2.X),
                Math.Min(p1.Y, p2.Y),
                Math.Abs(p1.X - p2.X),
                Math.Abs(p1.Y - p2.Y));
        }
    }
}

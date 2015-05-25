namespace VisioAutomation.UI
{
    internal class ColorPickerUtil
    {
        public static System.Drawing.Rectangle Rect(System.Drawing.RectangleF rect)
        {
            var new_rect = new System.Drawing.Rectangle();
            new_rect.X = (int) rect.X;
            new_rect.Y = (int) rect.Y;
            new_rect.Width = (int) rect.Width;
            new_rect.Height = (int) rect.Height;
            return new_rect;
        }

        public static System.Drawing.RectangleF Rect(System.Drawing.Rectangle rect)
        {
            var new_rect = new System.Drawing.RectangleF();
            new_rect.X = rect.X;
            new_rect.Y = rect.Y;
            new_rect.Width = rect.Width;
            new_rect.Height = rect.Height;
            return new_rect;
        }

        public static System.Drawing.Point Point(System.Drawing.PointF point)
        {
            return new System.Drawing.Point((int)point.X, (int)point.Y);
        }

        public static System.Drawing.PointF Center(System.Drawing.RectangleF rect)
        {
            var center = rect.Location;
            center.X += rect.Width/2;
            center.Y += rect.Height/2;
            return center;
        }

        public static void DrawFrame(System.Drawing.Graphics dc, System.Drawing.RectangleF r, float cornerRadius, System.Drawing.Color color)
        {
            var pen = new System.Drawing.Pen(color);
            if (cornerRadius <= 0)
            {
                dc.DrawRectangle(pen, ColorPickerUtil.Rect(r));
                return;
            }
            cornerRadius = (float)System.Math.Min(cornerRadius, System.Math.Floor(r.Width) - 2);
            cornerRadius = (float)System.Math.Min(cornerRadius, System.Math.Floor(r.Height) - 2);

            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(r.X, r.Y, cornerRadius, cornerRadius, 180, 90);
            path.AddArc(r.Right - cornerRadius, r.Y, cornerRadius, cornerRadius, 270, 90);
            path.AddArc(r.Right - cornerRadius, r.Bottom - cornerRadius, cornerRadius, cornerRadius, 0, 90);
            path.AddArc(r.X, r.Bottom - cornerRadius, cornerRadius, cornerRadius, 90, 90);
            path.CloseAllFigures();
            dc.DrawPath(pen, path);
        }

        public static void Draw2ColorBar(System.Drawing.Graphics dc, System.Drawing.RectangleF r, System.Windows.Forms.Orientation orientation, System.Drawing.Color c1, System.Drawing.Color c2)
        {
            var lr1 = r;
            float angle = 0;

            if (orientation == System.Windows.Forms.Orientation.Vertical)
            {angle = 270;}
            if (orientation == System.Windows.Forms.Orientation.Horizontal)
            {angle = 0;}

            if (lr1.Height > 0 && lr1.Width > 0)
            {
                using (var lb1 = new System.Drawing.Drawing2D.LinearGradientBrush(lr1, c1, c2, angle, false))
                {dc.FillRectangle(lb1, lr1);}
            }
        }

        public static void Draw3ColorBar(System.Drawing.Graphics dc, System.Drawing.RectangleF r, System.Windows.Forms.Orientation orientation, System.Drawing.Color c1, System.Drawing.Color c2,
                                         System.Drawing.Color c3)
        {
            // to draw a 3 color bar 2 gradient brushes are needed
            // one from c1 - c2 and c2 - c3
            var lr1 = r;
            var lr2 = r;
            float angle = 0;

            if (orientation == System.Windows.Forms.Orientation.Vertical)
            {
                angle = 270;

                lr1.Height = lr1.Height/2;
                lr2.Height = r.Height - lr1.Height;
                lr2.Y += lr1.Height;
            }
            if (orientation == System.Windows.Forms.Orientation.Horizontal)
            {
                angle = 0;

                lr1.Width = lr1.Width/2;
                lr2.Width = r.Width - lr1.Width;
                lr1.X = lr2.Right;
            }

            if (lr1.Height > 0 && lr1.Width > 0)
            {
                using (System.Drawing.Drawing2D.LinearGradientBrush lb2 = new System.Drawing.Drawing2D.LinearGradientBrush(lr2, c1, c2, angle, false),  lb1 = new System.Drawing.Drawing2D.LinearGradientBrush(lr1, c2, c3, angle, false) )
                {
                    dc.FillRectangle(lb1, lr1);
                    dc.FillRectangle(lb2, lr2);                   
                }
            }

            // with some sizes the first pixel in the gradient rectangle shows the opposite color
            // this is a workaround for that problem
            if (orientation == System.Windows.Forms.Orientation.Vertical)
            {
                using (System.Drawing.Pen pc2 = new System.Drawing.Pen(c2, 1), pc3 = new System.Drawing.Pen(c3, 1))
                {
                    dc.DrawLine(pc3, lr1.Left, lr1.Top, lr1.Right - 1, lr1.Top);
                    dc.DrawLine(pc2, lr2.Left, lr2.Top, lr2.Right - 1, lr2.Top);
                }
            }

            if (orientation == System.Windows.Forms.Orientation.Horizontal)
            {
                using (System.Drawing.Pen pc1 = new System.Drawing.Pen(c1, 1), pc2 = new System.Drawing.Pen(c2, 1), pc3 = new System.Drawing.Pen(c3, 1))
                {
                    dc.DrawLine(pc1, lr2.Left, lr2.Top, lr2.Left, lr2.Bottom - 1);
                    dc.DrawLine(pc2, lr2.Right, lr2.Top, lr2.Right, lr2.Bottom - 1);
                    dc.DrawLine(pc3, lr1.Right, lr1.Top, lr1.Right, lr1.Bottom - 1);
                }
            }
        }
    }
}
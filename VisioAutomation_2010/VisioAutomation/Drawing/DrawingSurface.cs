using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Drawing
{
    public struct DrawingSurface
    {
        public Microsoft.Office.Interop.Visio.Page Page;
        public Microsoft.Office.Interop.Visio.Master Master;
        public Microsoft.Office.Interop.Visio.Shape Shape;

        public DrawingSurface(IVisio.Page page)
        {
            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public DrawingSurface(IVisio.Master master)
        {
            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }


        public DrawingSurface(IVisio.Shape shape)
        {
            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }

        public IVisio.Shape DrawLine(VA.Drawing.Point p1, VA.Drawing.Point p2)
        {

            if (this.Master != null)
            {
                var shape = this.Master.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

        }

        public IVisio.Shape DrawPolyLine(IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawBezier(IList<VA.Drawing.Point> points, short degree, short flags)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
 
            if (this.Master != null)
            {
                var shape = this.Master.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawBezier(IList<VA.Drawing.Point> points)
        {
            short degree = 3;
            short flags = 0;
            var shape = this.DrawBezier(points, degree, flags);
            return shape;
        }

        public IVisio.Shape DrawOval(VA.Drawing.Rectangle rect)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawOval(VA.Drawing.Point center, double radius)
        {
            var A = center.Add(-radius, -radius);
            var B = center.Add(radius, radius);
            var rect = new VA.Drawing.Rectangle(A, B);

            return this.DrawOval(rect);
        }

        public IVisio.Shape DrawRectangle(VA.Drawing.Rectangle rect)
        {
            var shape = this.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawRectangle(x0,y0,x1,y1);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            if (this.Master != null)
            {
                var shape = this.Master.DrawLine(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawLine(x0, y0, x1, y1);

                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawLine(x0, y0, x1, y1);

                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawNURBS(IList<VA.Drawing.Point> controlpoints,
                                     IList<double> knots,
                                     IList<double> weights, int degree)
        {
            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = VA.Drawing.Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            if (this.Master != null)
            {
                var shape = this.Master.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

        }

    }
}
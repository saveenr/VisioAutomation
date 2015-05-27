using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Drawing
{

    public struct DrawingSurface
    {
        public readonly SurfaceTarget Target;

        public DrawingSurface(SurfaceTarget target)
        {
            this.Target = target;
        }

        public DrawingSurface(IVisio.Page page)
        {
         this.Target = new SurfaceTarget(page);
        }

        public DrawingSurface(IVisio.Master master)
        {
            this.Target = new SurfaceTarget(master);
        }


        public DrawingSurface(IVisio.Shape shape)
        {
            this.Target = new SurfaceTarget(shape);
        }

        public IVisio.Shape DrawLine(Point p1, Point p2)
        {

            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public IVisio.Shape DrawPolyLine(IList<Point> points)
        {
            var doubles_array = Point.ToDoubles(points).ToArray();

            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawPolyline(doubles_array, 0);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawPolyline(doubles_array, 0);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
        }

        public IVisio.Shape DrawBezier(IList<Point> points, short degree, short flags)
        {
            var doubles_array = Point.ToDoubles(points).ToArray();

            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawBezier(doubles_array, degree, flags);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawBezier(doubles_array, degree, flags);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public IVisio.Shape DrawBezier(IList<Point> points)
        {
            short degree = 3;
            short flags = 0;
            var shape = this.DrawBezier(points, degree, flags);
            return shape;
        }

        public IVisio.Shape DrawOval(Rectangle rect)
        {
            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
        }

        public IVisio.Shape DrawOval(Point center, double radius)
        {
            var A = center.Add(-radius, -radius);
            var B = center.Add(radius, radius);
            var rect = new Rectangle(A, B);

            return this.DrawOval(rect);
        }

        public IVisio.Shape DrawRectangle(Rectangle rect)
        {
            var shape = this.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawLine(x0, y0, x1, y1);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawLine(x0, y0, x1, y1);

                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawLine(x0, y0, x1, y1);

                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawNURBS(IList<Point> controlpoints,
            IList<double> knots,
            IList<double> weights, int degree)
        {
            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            if (this.Target.Master != null)
            {
                var shape = this.Target.Master.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Target.Page != null)
            {
                var shape = this.Target.Page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Target.Shape != null)
            {
                var shape = this.Target.Shape.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");

        }

        public short[] DropManyU(
            IList<IVisio.Master> masters,
            IEnumerable<Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(masters));
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            if (this.Target.Master != null)
            {
                this.Target.Master.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Target.Page != null)
            {
                this.Target.Page.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Target.Shape != null)
            {
                this.Target.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            short[] outids = (short[]) outids_sa;
            return outids;
        }

        public IVisio.Shape Drop(
            IVisio.Master master,
            Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            if (this.Target.Master != null)
            {
                return this.Target.Master.Drop(master, point.X, point.Y);
            }
            else if (this.Target.Page != null)
            {
                return this.Target.Page.Drop(master, point.X, point.Y);
            }
            else if (this.Target.Shape != null)
            {
                return this.Target.Shape.Drop(master, point.X, point.Y);
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public IVisio.Shape DrawQuarterArc(Point p0, Point p1, IVisio.VisArcSweepFlags flags)
        {
            if (this.Target.Master != null)
            {
                return this.Target.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Target.Page != null)
            {
                return this.Target.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Target.Shape != null)
            {
                return this.Target.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }

            throw new System.ArgumentException("Unhandled Drawing Surface");
            
        }

        public Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            if (this.Target.Master != null)
            {
                this.Target.Master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Target.Page != null)
            {
                this.Target.Page.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Target.Shape != null)
            {
                this.Target.Shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var r = new Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

    }
}


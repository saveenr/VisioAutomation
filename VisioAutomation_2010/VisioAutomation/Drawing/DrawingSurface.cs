using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Drawing
{

    public struct DrawingSurface
    {
        public readonly SurfaceTarget Target;

        public DrawingSurface(SurfaceTarget k)
        {
            this.Target = k;
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

        public IVisio.Master Master
        {
            get
            {
                return this.Target.Master;
            }
        }

        public IVisio.Page Page
        {
            get
            {
                return this.Target.Page;
            }
        }

        public IVisio.Shape Shape
        {
            get
            {
                return this.Target.Shape;
            }
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
                var shape = this.Master.DrawRectangle(x0, y0, x1, y1);
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
                var shape = this.Master.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Page != null)
            {
                var shape = this.Page.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else if (this.Shape != null)
            {
                var shape = this.Shape.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
                return shape;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

        }

        public short[] DropManyU(
            IList<IVisio.Master> masters,
            IEnumerable<VA.Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException("masters");
            }

            if (masters.Count < 1)
            {
                return new short[0];
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = VA.Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            if (this.Master != null)
            {
                this.Master.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Page != null)
            {
                this.Page.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Shape != null)
            {
                this.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa);
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
            VA.Drawing.Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            if (this.Master != null)
            {
                return this.Master.Drop(master, point.X, point.Y);
            }
            else if (this.Page != null)
            {
                return this.Page.Drop(master, point.X, point.Y);
            }
            else if (this.Shape != null)
            {
                return this.Shape.Drop(master, point.X, point.Y);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }

        public IVisio.Shape DrawQuarterArc(VA.Drawing.Point p0, VA.Drawing.Point p1, IVisio.VisArcSweepFlags flags)
        {
            if (this.Master != null)
            {
                return this.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Page != null)
            {
                return this.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Shape != null)
            {
                return this.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }
        }



        public VA.Drawing.Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            if (this.Master != null)
            {
                this.Master.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Page != null)
            {
                this.Page.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Shape != null)
            {
                this.Shape.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var r = new VA.Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public IVisio.Shapes Shapes
        {
            get
            {

                IVisio.Shapes shapes;

                if (this.Master != null)
                {

                    shapes = this.Master.Shapes;
                }
                else if (this.Page != null)
                {
                    shapes = this.Page.Shapes;
                }
                else if (this.Shape != null)
                {
                    shapes = this.Shape.Shapes;
                }
                else
                {
                    throw new System.ArgumentException("Unhandled Drawing Surface");
                }
                return shapes;
            }

        }

        public List<IVisio.Shape> GetAllShapes()
        {
            IVisio.Shapes shapes;

            if (this.Master != null)
            {

                shapes = this.Master.Shapes;
            }
            else if (this.Page != null)
            {
                shapes = this.Page.Shapes;
            }
            else if (this.Shape != null)
            {
                shapes = this.Shape.Shapes;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var list = new List<IVisio.Shape>();
            list.AddRange(shapes.AsEnumerable());

            return list;
        }

        public VA.ShapeSheet.ShapeSheetSurface ToShapeSheetSurface()
        {
            return new ShapeSheetSurface(this.Target);
        }
    }
}


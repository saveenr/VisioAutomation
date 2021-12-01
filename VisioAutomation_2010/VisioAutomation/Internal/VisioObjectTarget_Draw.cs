using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Internal
{
    internal readonly partial struct VisioObjectTarget
    {
        public IVisio.Shape DrawOval(Core.Rectangle rect)
        {
            IVisio.Shape shape;

            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawRectangle(Core.Rectangle rect)
        {
            IVisio.Shape shape;

            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawBezier(IList<Core.Point> points)
        {
            var doubles_array = VisioAutomation.Core.Point.ToDoubles(points).ToArray();
            short degree = 3;
            short flags = 0;

            IVisio.Shape shape;

            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawBezier(doubles_array, degree, flags);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawBezier(doubles_array, degree, flags);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawBezier(doubles_array, degree, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawPolyline(IList<Core.Point> points)
        {
            var doubles_array = Core.Point.ToDoubles(points).ToArray();
            IVisio.Shape shape;

            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawPolyline(doubles_array, 0);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawPolyline(doubles_array, 0);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawPolyline(doubles_array, 0);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }


        public IVisio.Shape DrawQuarterArc(
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            IVisio.Shape shape;
            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawLine(
            Core.Point p0,
            Core.Point p1)
        {
            IVisio.Shape shape;
            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawNurbs(
            IList<Core.Point> controlpoints,
            IList<double> knots,
            IList<double> weights,
            int degree)
        {
            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = Core.Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            IVisio.Shape shape;
            if (this.Category == VisioObjectCategory.Shape)
            {
                shape = this.Shape.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                shape = this.Master.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                shape = this.Page.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }


        public Core.Rectangle GetBoundingBox(IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;


            if (this.Category == VisioObjectCategory.Shape)
            {
                this.Shape.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                this.Master.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                this.Page.BoundingBox((short) args, out bbx0, out bby0, out bbx1, out bby1);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var r = new Core.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }
    }
}
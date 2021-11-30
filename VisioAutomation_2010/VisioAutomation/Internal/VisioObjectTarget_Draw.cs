using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Internal
{
    internal readonly partial struct VisioObjectTarget
    {

        public Microsoft.Office.Interop.Visio.Shape DrawOval(Core.Rectangle rect)
        {

            var visobjtarget = this;

            IVisio.Shape shape;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawRectangle(Core.Rectangle rect)
        {
            var visobjtarget = this;


            IVisio.Shape shape;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public  IVisio.Shape DrawBezier(IList<Core.Point> points)
        {

            var visobjtarget = this;


            var doubles_array = VisioAutomation.Core.Point.ToDoubles(points).ToArray();
            short degree = 3;
            short flags = 0;

            IVisio.Shape shape;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawBezier(doubles_array, degree, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawBezier(doubles_array, degree, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawBezier(doubles_array, degree, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }

        public IVisio.Shape DrawPolyline(IList<Core.Point> points)
        {

            var visobjtarget = this;

            var doubles_array = Core.Point.ToDoubles(points).ToArray();
            IVisio.Shape shape;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawPolyline(doubles_array, 0);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawPolyline(doubles_array, 0);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawPolyline(doubles_array, 0);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }



        public  IVisio.Shape DrawQuarterArc(
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {

            var visobjtarget = this;


            IVisio.Shape shape;
            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
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

            var visobjtarget = this;

            IVisio.Shape shape;
            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawLine(p0.X, p0.Y, p1.X, p1.Y);
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

            var visobjtarget = this;


            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var flags = 0;
            double[] pts_dbl_a = Core.Point.ToDoubles(controlpoints).ToArray();
            double[] kts_dbl_a = knots.ToArray();
            double[] weights_dbl_a = weights.ToArray();

            IVisio.Shape shape;
            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                shape = visobjtarget.Shape.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                shape = visobjtarget.Master.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                shape = visobjtarget.Page.DrawNURBS((short)degree, (short)flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return shape;
        }
    }
}
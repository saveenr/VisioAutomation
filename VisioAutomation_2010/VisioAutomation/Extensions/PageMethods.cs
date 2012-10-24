using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA = VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static partial class PageMethods
    {
        public static void Activate(this IVisio.Page page)
        {
            VA.Pages.PageHelper.Activate(page);
        }

        public static void ResizeToFitContents(this IVisio.Page page, double borderwidth, double borderheight)
        {
            var bordersize = new VA.Drawing.Size(borderwidth, borderheight);
            VA.Pages.PageHelper.ResizeToFitContents(page, bordersize);
        }

        public static void ResizeToFitContents(this IVisio.Page page, VA.Drawing.Size bordersize)
        {
            VA.Pages.PageHelper.ResizeToFitContents(page, bordersize);
        }

        public static VA.Drawing.Size GetSize(this IVisio.Page page)
        {
            return VA.Pages.PageHelper.GetSize(page);
        }

        public static void SetSize(this IVisio.Page page, VA.Drawing.Size size)
        {
            VA.Pages.PageHelper.SetSize(page, size);
        }

        public static void SetSize(this IVisio.Page page, double x, double y)
        {
            VA.Pages.PageHelper.SetSize(page, new VA.Drawing.Size(x, y));
        }

        public static IVisio.Shape DrawLine(this IVisio.Page page, VA.Drawing.Point p1, VA.Drawing.Point p2)
        {
            var shape = page.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static IVisio.Shape DrawOval(this IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var shape = page.DrawOval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, VA.Drawing.Rectangle rect)
        {
            var shape = page.DrawRectangle(rect.Left, rect.Bottom, rect.Right, rect.Top);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
            var shape = page.DrawBezier(doubles_array, 3, 0);
            return shape;
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<VA.Drawing.Point> points)
        {
            var doubles_array = VA.Drawing.Point.ToDoubles(points).ToArray();
            var shape = page.DrawPolyline(doubles_array, 0);
            return shape;
        }

        public static IVisio.Shape DrawNURBS(this IVisio.Page page, IList<VA.Drawing.Point> controlpoints,
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

            var shape = page.DrawNURBS((short) degree, (short) flags, pts_dbl_a, kts_dbl_a, weights_dbl_a);
            return shape;
        }

        public static IVisio.Shape Drop(
            this IVisio.Page page,
            IVisio.Master master,
            VA.Drawing.Point point)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            return page.Drop(master, point.X, point.Y);
        }

        public static short[] DropManyU(
            this IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<VA.Drawing.Point> points)
        {
            short[] shapeids = VA.Pages.PageHelper.DropManyU(page, masters, points);
            return shapeids;
        }
    }
}
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class PageMethods
    {
        public static void ResizeToFitContents(this IVisio.Page page, Geometry.Size padding)
        {
            // first perform the native resizetofit
            page.ResizeToFitContents();

            if ((padding.Width > 0.0) || (padding.Height > 0.0))
            {
                // if there is any additional padding requested
                // we need to further handle the page

                // first determine the desired page size including the padding
                // and set the new size

                var old_size = VisioAutomation.Pages.PageHelper.GetSize(page);
                var new_size = old_size + padding.Multiply(2, 2);
                VisioAutomation.Pages.PageHelper.SetSize(page, new_size);

                // The page has the correct size, but
                // the contents will be offset from the correct location
                page.CenterDrawing();
            }
        }

        public static IVisio.Shape DrawLine(this IVisio.Page page, Geometry.Point p1, Geometry.Point p2)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawLine(p1,p2);
            return shape;
        }

        public static IVisio.Shape DrawOval(this IVisio.Page page, Geometry.Rectangle rect)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawOval(rect);
            return shape;
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, Geometry.Rectangle rect)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawRectangle(rect);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Geometry.Point> points)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawBezier(points);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Geometry.Point> points, short degree, short flags)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawBezier(points, degree, flags);
            return shape;
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<Geometry.Point> points)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawBezier(points);
            return shape;
        }

        public static IVisio.Shape DrawNURBS(
            this IVisio.Page page, 
            IList<Geometry.Point> controlpoints,
            IList<double> knots,
            IList<double> weights, 
            int degree)
        {
            var surface = new SurfaceTarget(page);
            var shape = surface.DrawNURBS(controlpoints, knots, weights, degree);
            return shape;
        }

        public static IVisio.Shape Drop(
            this IVisio.Page page,
            IVisio.Master master,
            Geometry.Point point)
        {
            var surface = new SurfaceTarget(page);
            return surface.Drop(master, point);
        }

        public static short[] DropManyU(
            this IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<Geometry.Point> points)
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
            var xy_array = Geometry.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

        public static IEnumerable<IVisio.Page> ToEnumerable(this IVisio.Pages pages)
        {
            return ExtensionHelpers.ToEnumerable(() => pages.Count, i => pages[i + 1]);
        }

        public static List<IVisio.Page> ToList(this IVisio.Pages pages)
        {
            return ExtensionHelpers.ToList(() => pages.Count, i => pages[i + 1]);
        }

        public static string[] GetNamesU(this IVisio.Pages pages)
        {
            System.Array names_sa;
            pages.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}
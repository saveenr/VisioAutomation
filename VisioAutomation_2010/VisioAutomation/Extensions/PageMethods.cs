using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Drawing;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class PageMethods
    {
        public static void ResizeToFitContents(this IVisio.Page page, Drawing.Size padding)
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

        public static IVisio.Shape DrawLine(this IVisio.Page page, Drawing.Point p1, Drawing.Point p2)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static IVisio.Shape DrawOval(this IVisio.Page page, Drawing.Rectangle rect)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawOval(rect);
            return shape;
        }

        public static IVisio.Shape DrawRectangle(this IVisio.Page page, Drawing.Rectangle rect)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawRectangle(rect);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Drawing.Point> points)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawBezier(points);
            return shape;
        }

        public static IVisio.Shape DrawBezier(this IVisio.Page page, IList<Drawing.Point> points, short degree, short flags)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawBezier(points, degree, flags);
            return shape;
        }

        public static IVisio.Shape DrawPolyline(this IVisio.Page page, IList<Drawing.Point> points)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawBezier(points);
            return shape;
        }

        public static IVisio.Shape DrawNURBS(
            this IVisio.Page page, 
            IList<Drawing.Point> controlpoints,
            IList<double> knots,
            IList<double> weights, int degree)
        {
            var surface = new Drawing.DrawingSurface(page);
            var shape = surface.DrawNURBS(controlpoints, knots, weights, degree);
            return shape;
        }

        public static IVisio.Shape Drop(
            this IVisio.Page page,
            IVisio.Master master,
            Drawing.Point point)
        {
            var surface = new Drawing.DrawingSurface(page);
            return surface.Drop(master, point);
        }

        public static short[] DropManyU(
            this IVisio.Page page,
            IList<IVisio.Master> masters,
            IEnumerable<Drawing.Point> points)
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
            var xy_array = Drawing.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            page.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }

        public static short[] DropManyU(this IVisio.Page page, IList<IVisio.Master> masters, IEnumerable<Point> points, IList<string> names)
   	    {
            var surface = new VisioAutomation.Drawing.DrawingSurface(page);
            short[] shapeids = surface.DropManyU(masters, points);

            if (names != null)
            {
                var page_shapes = page.Shapes;
                //ShapeIDs should come back in the same order... if Names for the masters are passed in, assign the Name. Makes it easier to find later
                for (int i = 0; i < shapeids.Length; i++)
                {
                    var shapeid = shapeids[i];
                    var cur_shape = page_shapes[shapeid];
                    cur_shape.Name = names[i];
                }
            }
            return shapeids;
        }

        public static IEnumerable<IVisio.Page> ToEnumerable(this IVisio.Pages pages)
        {
            short count = pages.Count;
            for (int i = 0; i < count; i++)
            {
                yield return pages[i + 1];
            }
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
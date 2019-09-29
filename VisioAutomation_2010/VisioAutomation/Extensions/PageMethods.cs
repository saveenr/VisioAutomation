using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods
    {
        public static void ResizeToFitContents(this IVisio.Page page, Geometry.Size padding)
        {
            Pages.PageHelper.ResizeToFitContents(page, padding);
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

        public static IVisio.Shape DrawNurbs(
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
            return Pages.PageHelper.DropManyU(page, masters, points);
        }

        public static IEnumerable<IVisio.Page> ToEnumerable(this IVisio.Pages pages)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => pages.Count, i => pages[i + 1]);
        }

        public static List<IVisio.Page> ToList(this IVisio.Pages pages)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => pages.Count, i => pages[i + 1]);
        }

        public static string[] GetNamesU(this IVisio.Pages pages)
        {
            System.Array names_sa;
            pages.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }

        public static string[] GetFormulasU(this IVisio.Page page, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            page.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = SurfaceTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this IVisio.Page page, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {

            var flags = SurfaceTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            page.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = SurfaceTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }

    }
}
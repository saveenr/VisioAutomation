using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods
    {
        public static IVisio.Shape DrawLine(
            this IVisio.Shape shape,
            Geometry.Point p1, Geometry.Point p2)
        {
            var surface = new SurfaceTarget(shape);
            var s = surface.DrawLine(p1, p2);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Shape shape,
            Geometry.Point p0,
            Geometry.Point p1, 
            IVisio.VisArcSweepFlags flags)
        {
            var surface = new SurfaceTarget(shape);
            var s = surface.DrawQuarterArc(p0, p1, flags);
            return s;
        }

        public static Geometry.Rectangle GetBoundingBox(
            this IVisio.Shape shape, 
            IVisio.VisBoundingBoxArgs args)
        {
            var surface = new SurfaceTarget(shape);
            var r = surface.GetBoundingBox(args);
            return r;
        }

        public static Geometry.Point XYFromPage(
            this IVisio.Shape shape,
            Geometry.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new Geometry.Point(xprime, yprime);
        }

        public static Geometry.Point XYToPage(
            this IVisio.Shape shape,
            Geometry.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new Geometry.Point(xprime, yprime);
        }

        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Shapes shapes)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => shapes.Count, i => shapes[i + 1]);
        }

        public static List<IVisio.Shape> ToList(this IVisio.Shapes shapes)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => shapes.Count, i => shapes[i + 1]);
        }

        public static string[] GetFormulasU(this IVisio.Shape shape, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            shape.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = SurfaceTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this IVisio.Shape shape, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            var flags = SurfaceTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = SurfaceTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }


    }
}
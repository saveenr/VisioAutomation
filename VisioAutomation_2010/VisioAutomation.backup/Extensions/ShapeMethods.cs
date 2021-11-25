﻿using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods
    {
        public static Geometry.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            shape.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new VisioAutomation.Geometry.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IVisio.Shape DrawLine(
            this IVisio.Shape shape,
            Geometry.Point p1, Geometry.Point p2)
        {
            var s = shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return s;
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Shape shape,
            Geometry.Point p0,
            Geometry.Point p1, 
            IVisio.VisArcSweepFlags flags)
        {
            var s = shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            return s;
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
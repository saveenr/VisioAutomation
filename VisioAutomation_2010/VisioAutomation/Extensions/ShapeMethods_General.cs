﻿using System.Collections.Generic;
using VisioAutomation.Internal;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_General
    {
        public static IEnumerable<IVisio.Shape> ToEnumerable(this IVisio.Shapes shapes)
        {
            return CollectionHelpers.ToEnumerable(() => shapes.Count, i => shapes[i + 1]);
        }

        public static List<IVisio.Shape> ToList(this IVisio.Shapes shapes)
        {
            return CollectionHelpers.ToList(() => shapes.Count, i => shapes[i + 1]);
        }

        public static Core.Rectangle GetBoundingBox(this IVisio.Shape shape, IVisio.VisBoundingBoxArgs args)
        {
            return shape.Wrap().GetBoundingBox(args);
        }


        public static Core.Point XYFromPage(
            this IVisio.Shape shape,
            Core.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff767213.aspx
            double xprime;
            double yprime;
            shape.XYFromPage(xy.X, xy.Y, out xprime, out yprime);
            return new Core.Point(xprime, yprime);
        }

        public static Core.Point XYToPage(
            this IVisio.Shape shape,
            Core.Point xy)
        {
            // MSDN: http://msdn.microsoft.com/en-us/library/office/ff766239.aspx
            double xprime;
            double yprime;
            shape.XYToPage(xy.X, xy.Y, out xprime, out yprime);
            return new Core.Point(xprime, yprime);
        }

        internal static VisioAutomation.Internal.VisioObjectTarget Wrap(this IVisio.Shape shape)
        {
            return new VisioAutomation.Internal.VisioObjectTarget(shape);
        }
    }
}
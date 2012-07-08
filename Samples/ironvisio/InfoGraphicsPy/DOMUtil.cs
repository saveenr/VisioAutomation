using System;
using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    internal static class DOMUtil
    {
        public static List<VA.DOM.BaseShape> DrawOvals(VA.DOM.ShapeCollection dom, IList<VA.Drawing.Rectangle> rects)
        {
            var dom_shapes = new List<VA.DOM.BaseShape>();
            foreach (var rect in rects)
            {
                var dom_shape = dom.DrawOval(rect);
                dom_shape.Cells.Width = rect.Width;
                dom_shape.Cells.Height = rect.Height;
                dom_shapes.Add(dom_shape);
            }

            return dom_shapes;
        }

        public static List<VA.DOM.Shape> DrawRects(VA.DOM.ShapeCollection dom, IList<VA.Drawing.Rectangle> rects, IVisio.Master rectmaster)
        {
            var dom_shapes = new List<VA.DOM.Shape>();
            foreach (var rect in rects)
            {
                var dom_shape = dom.Drop(rectmaster, rect.Center);
                dom_shape.Cells.Width = rect.Width;
                dom_shape.Cells.Height = rect.Height;
                dom_shapes.Add(dom_shape);
            }

            return dom_shapes;
        }

        public static List<IVisio.Shape> DrawRects(IList<VA.Drawing.Rectangle> rects, IVisio.Master rectmaster, IVisio.Page page)
        {
            var dom_shapes = new List<VA.DOM.Shape>();
            var dom = new VA.DOM.ShapeCollection();
            foreach (var rect in rects)
            {
                var dom_shape = dom.Drop(rectmaster, rect.Center);
                dom_shape.Cells.Width = rect.Width;
                dom_shape.Cells.Height = rect.Height;
                dom_shapes.Add(dom_shape);
            }

            dom.Render(page);

            var shapes = new List<IVisio.Shape>();
            foreach (var dom_shape in dom_shapes)
            {
                shapes.Add(dom_shape.VisioShape);
            }

            return shapes;
        }

        public static List<double> ConstructPositions(int numcols, double width, double sep)
        {
            var iwidths = new List<double>();
            for (int i = 0; i < numcols; i++)
            {
                iwidths.Add(width);
            }
            var widths = ConstructPositions(iwidths, sep);
            return widths;
        }

        public static List<double> ConstructPositions(IList<double> iwidths, double sep)
        {
            int numcols = iwidths.Count;
            var widths = new List<double>();

            for (int i = 0; i < numcols; i++)
            {
                widths.Add(iwidths[i]);
                if (i < numcols - 1)
                {
                    widths.Add(sep);
                }
            }
            return widths;
        }

    }
}

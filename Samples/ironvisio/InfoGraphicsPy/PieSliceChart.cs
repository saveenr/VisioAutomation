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
    public class PieSliceChart: Chart
    {
        public void Draw(IVisio.Page page, IVisio.Master rectmaster )
        {
            double cellwidth = 0.5;
            double hsep = 0.10;
            double vsep = 0.10;
            double cellheight = cellwidth;
            double catheight = 0.5;
            var cats = new[] { "A", "B", "C", "D", "E" };
            var datapoints = new DataPoints(new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 });
            var normalized_values = datapoints.GetNormalizedValues();
            var widths = ConstructPositions(datapoints.Count(), cellwidth, hsep);
            var heights = ConstructPositions(new[] { catheight, cellheight }, vsep);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var circle_shapes = new List<VA.DOM.Oval>();
            var slice_shapes = new List<VA.DOM.PieSlice>();
            for (int i = 0; i < datapoints.Count; i++)
            {
                var dp = datapoints[i];
                double start = 0;
                double end = 360*normalized_values[i];
                double radius = top_rects[i].Width/2.0;

                var circle_shape = dom.DrawOval(top_rects[i]);
                circle_shapes.Add(circle_shape);

                var dom_shape = dom.DrawPieSlice(top_rects[i].Center, radius, start, end);
                slice_shapes.Add(dom_shape);
            }
            var cat_shapes = this.DrawRects(dom, cat_rects, rectmaster);

            for (int i = 0; i < datapoints.Count; i++)
            {
                slice_shapes[i].Text = datapoints[i].Text.ToString();
                cat_shapes[i].Text = cats[i];
            }

            foreach (var shape in circle_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = "rgb(255,255,255)";
                cells.LineColor = "rgb(220,220,220)";

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = "rgb(240,240,240)";
                cells.LineColor = "rgb(220,220,220)";

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = "0";
                cells.LineWeight = "0.0";
                cells.LinePattern = "0";
            }
            dom.Render(page);


        }
    }
}

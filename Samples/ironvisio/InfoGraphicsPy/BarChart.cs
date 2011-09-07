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
    public class BarChart : Chart
    {
        private double cellwidth = 0.5;
        public double HorizontalSeparation = 0.10;
        public double VerticalSeparation = 0.10;
        public double CellHeight = 0.5;
        public double CategoryLabelHeight = 0.5;

        public string[] CategoryLabels;
        public DataPoints DataPoints;

        public BarChart(DataPoints dps, string[] cats)
        {
            this.DataPoints = dps;
            this.CategoryLabels = cats;

        }

        public double CellWidth
        {
            get { return cellwidth; }
            set { cellwidth = value; }
        }

        public void Draw(Session session)
        {
            double cellwidth = 0.5;
            double hsep = 0.10;
            double vsep = 0.10;
            double cellheight = 4;
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

            var bar_rects = new List<VA.Drawing.Rectangle>(top_rects.Count);
            for (int i = 0; i < top_rects.Count; i++)
            {
                var r = top_rects[i];
                var size = new VA.Drawing.Size(r.Width, normalized_values[i] * cellheight);
                var bar_rect = new VA.Drawing.Rectangle(r.LowerLeft, size);
                bar_rects.Add(bar_rect);
            }
            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var bar_shapes = this.DrawRects(dom, bar_rects, session.MasterRectangle);
            var cat_shapes = this.DrawRects(dom, cat_rects, session.MasterRectangle);

            for (int i = 0; i < datapoints.Count; i++)
            {
                bar_shapes[i].Text = datapoints[i].Text.ToString();
                cat_shapes[i].Text = cats[i];
            }

            foreach (var shape in bar_shapes)
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
            dom.Render(session.Page);
        }
    }

}

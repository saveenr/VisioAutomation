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
    public class DoughnutSliceChart : GridChart
    {
        public DoughnutSliceChart(DataPoints dps, string[] cats) :
            base(dps,cats)
        {
        }

        public void Draw(Session session)
        {
            var normalized_values = DataPoints.GetNormalizedValues();
            var widths = DOMUtil.ConstructPositions(DataPoints.Count(), this.CellWidth, HorizontalSeparation);
            var heights = DOMUtil.ConstructPositions(new[] { CategoryLabelHeight, CellHeight }, VerticalSeparation);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            var circle_shapes = new List<VA.DOM.Arc>();
            var slice_shapes = new List<VA.DOM.Arc>();
            for (int i = 0; i < DataPoints.Count; i++)
            {
                var dp = DataPoints[i];
                var start = 0.0;
                var end = System.Math.PI * 2.0 * normalized_values[i];
                double radius = top_rects[i].Width / 2.0;

                var circle_shape = dom.DrawArc(top_rects[i].Center, radius * 0.7, radius, start, System.Math.PI * 2.0);
                circle_shapes.Add(circle_shape);

                var dom_shape = dom.DrawArc(top_rects[i].Center, radius*0.7, radius, start, end);
                slice_shapes.Add(dom_shape);
            }

            var cat_shapes = DOMUtil.DrawRects(dom, cat_rects, session.MasterRectangle);

            for (int i = 0; i < DataPoints.Count; i++)
            {
                slice_shapes[i].Text = new VA.Text.Markup.TextElement(DataPoints[i].Text);
                cat_shapes[i].Text = new VA.Text.Markup.TextElement(CategoryLabels[i]);
            }

            foreach (var shape in circle_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = NonValueColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = ValueFillColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = CategoryFillPattern;
                cells.LineWeight = CategoryLineWeight;
                cells.LinePattern = CategoryLinePattern;
            }

            dom.Render(session.Page);
        }
    }

}

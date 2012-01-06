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
    public class VerticalBarChart : GridChart
    {
        public VerticalBarChart(DataPoints dps, string[] cats) :
            base(dps,cats)
        {
        }

        public void Draw(Session session)
        {
            var normalized_values = this.DataPoints.GetNormalizedValues();

            var widths = DOMUtil.ConstructPositions(this.DataPoints.Count, CellWidth, this.HorizontalSeparation);
            var heights = DOMUtil.ConstructPositions(new[] { this.CategoryLabelHeight, this.CellHeight }, this.VerticalSeparation);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var bar_rects = new List<VA.Drawing.Rectangle>(top_rects.Count);
            for (int i = 0; i < top_rects.Count; i++)
            {
                var r = top_rects[i];
                var size = new VA.Drawing.Size(r.Width, normalized_values[i] * this.CellHeight);
                var bar_rect = new VA.Drawing.Rectangle(r.LowerLeft, size);
                bar_rects.Add(bar_rect);
            }
            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            var bar_shapes = DOMUtil.DrawRects(dom, bar_rects, session.MasterRectangle);
            var cat_shapes = DOMUtil.DrawRects(dom, cat_rects, session.MasterRectangle);

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                bar_shapes[i].Text = new VA.Text.Markup.TextElement( this.DataPoints[i].Text);
                cat_shapes[i].Text = new VA.Text.Markup.TextElement(this.CategoryLabels[i]);
            }

            foreach (var shape in bar_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = this.ValueFillColor;
                cells.LineColor = this.LineLightBorder;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = this.CategoryFillPattern;
                cells.LineWeight = this.CategoryLineWeight;
                cells.LinePattern = this.CategoryLinePattern;
            }
            dom.Render(session.Page);
        }
    }
}

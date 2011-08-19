using System;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class HorizontalBarChart : BaseBarChart
    {
        double MaxBarWidth = 5.0;
        double label_width = 2.0;

        public HorizontalBarChart()
        {
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            //var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(rc.PageWidth, TileHeight));

            var ul = rc.CurrentUpperLeft.Add(margin);

            var xdoc = new VA.DOM.Document();

            VA.Drawing.Rectangle bb = new VA.Drawing.Rectangle(ul, new VA.Drawing.Size(0,0));

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];

                double max_left = ul.X + margin.Width + label_width; // need to account for labels
                double top = ul.Y - margin.Height - bar_thickness;

                // calc bar rect
                double bar_width = dp.Value / maxval * MaxBarWidth;

                double skip = (bar_thickness + _barDistance);
                var bar_ll = new VA.Drawing.Point(max_left, top).Add(1, -i * skip);
                var bar_ur = bar_ll.Add(bar_width, bar_thickness);

                var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);

                // draw bar shape
                var bar_shape = xdoc.DrawRectangle(bar_rect);
                bar_shape.Text = dp.Value.ToString();

                FormatBarShape(rc, bar_shape);

                var label_ll = bar_rect.LowerLeft.Subtract(margin.Width, 0).Add(-0.5, 0).Add(-label_width, 0);
                var label_ur = label_ll.Add(label_width, bar_thickness);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);

                var label_shape = xdoc.DrawRectangle(label_rect);
                FormatLabelShape(rc, dp, label_shape);

                bb = VA.Drawing.DrawingUtil.GetBoundingBox(bb, label_rect);
                bb = VA.Drawing.DrawingUtil.GetBoundingBox(bb, bar_rect);
            }

            xdoc.Render(rc.Page);


            return new VA.Drawing.Size(rc.PageWidth, bb.Height);

        }


        private static void FormatLabelShape(RenderContext rc, DataPoint dp, DOM.Rectangle label_shape)
        {
            label_shape.Text = dp.Value.ToString();
            label_shape.ShapeCells.LinePattern = 0;
            label_shape.ShapeCells.LineWeight = 0.0;
            label_shape.ShapeCells.FillPattern = 0;
            label_shape.ShapeCells.VerticalAlign = 0;
            label_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);
            label_shape.Text = dp.Label;
        }

        private void FormatBarShape(RenderContext rc, DOM.Rectangle bar_shape)
        {
            bar_shape.ShapeCells.LinePattern = 0;
            bar_shape.ShapeCells.LineWeight = 0.0;
            bar_shape.ShapeCells.FillForegnd = ValueColor.ToFormula();
            bar_shape.ShapeCells.VerticalAlign = 0;
            bar_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);
            bar_shape.ShapeCells.VerticalAlign = 1;
            bar_shape.ShapeCells.HAlign = 1;
        }

    }
}

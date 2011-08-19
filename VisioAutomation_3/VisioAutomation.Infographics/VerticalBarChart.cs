using System;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class VerticalBarChart : BaseBarChart
    {
        public VerticalBarChart()
        {
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(rc.PageWidth,TileHeight));
            var inner_ll = bkrect.LowerLeft.Add(margin);
            var inner_ur = bkrect.UpperRight.Subtract(margin);
            var innerrect = new VA.Drawing.Rectangle(inner_ll, inner_ur);

            var bararea_ll = innerrect.LowerLeft.Add(0, _labelHeight);
            var bararea_ur = innerrect.UpperRight;
            var bararea_rect = new VA.Drawing.Rectangle(bararea_ll, bararea_ur);
            
            var xdoc = new VA.DOM.Document();
           

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];

                double startpos = 0;
                startpos = bkrect.LowerLeft.Y;

                var bar_rect = GetBarRect(i, startpos, maxval, bararea_rect, dp);

                var bar_shape = xdoc.DrawRectangle(bar_rect);
                bar_shape.Text = dp.Value.ToString();

                FormatBarShape(rc, bar_shape);

                var label_rect = GetLabelRect(bar_rect);

                var label_shape = xdoc.DrawRectangle(label_rect);
                FormatLabelShape(rc, dp, label_shape);
            }

            xdoc.Render(rc.Page);

            return bkrect.Size;

        }

        private Rectangle GetLabelRect(Rectangle bar_rect)
        {
                var label_ll = bar_rect.LowerLeft.Subtract(0, margin.Height).Add(0, -0.5);
                var label_ur = label_ll.Add(bar_thickness, _labelHeight);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);
                return label_rect;
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

                bar_shape.ShapeCells.VerticalAlign = 0;
                bar_shape.ShapeCells.HAlign = 0;
        }

        private Rectangle GetBarRect(int i, double start_pos, double maxval, Rectangle bararea_rect, DataPoint dp)
        {
            double bar_length = dp.Value / maxval * bararea_rect.Height;

            var bar_ll = new VA.Drawing.Point(bar_thickness + _barDistance, start_pos).Multiply(i, 1).Add(margin.Width, margin.Height + _labelHeight);
            var bar_ur = bar_ll.Add(bar_thickness, bar_length);

            var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);
            return bar_rect;
        }
    }


}

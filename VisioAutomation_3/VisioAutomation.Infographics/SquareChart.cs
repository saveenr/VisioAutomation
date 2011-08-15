using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class SquareChart : Block
    {
        public IList<DataPoint> DataPoints;
        public VA.Drawing.ColorRGB ValueColor = new VA.Drawing.ColorRGB(0xa0a0a0);
        public VA.Drawing.ColorRGB NonValueColor = new VA.Drawing.ColorRGB(0xffffff);
        double TileHeight = 3.0;

        public SquareChart()
        {
            this.DataPoints = new List<DataPoint>();
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {

            var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(rc.PageWidth,TileHeight));

            double bar_distance = 0.0125;
            double lower_y = bkrect.LowerLeft.Y;
            double maxval = 180.0;
            double label_height = 0.5;

            var margin = new VA.Drawing.Size(0.25, 0.25);
            var inner_ll = bkrect.LowerLeft.Add(margin);
            var inner_ur = bkrect.UpperRight.Subtract(margin);
            var innerrect = new VA.Drawing.Rectangle(inner_ll, inner_ur);

            var bararea_ll = innerrect.LowerLeft.Add(0, label_height);
            var bararea_ur = innerrect.UpperRight;
            var bararea_rect = new VA.Drawing.Rectangle(bararea_ll, bararea_ur);
            
            var xdoc = new VA.DOM.Document();

            var tilerect = xdoc.DrawRectangle(bkrect);
            tilerect.ShapeCells.FillForegnd = rc.TileReal.ToFormula();
            tilerect.ShapeCells.LineWeight = 0;
            tilerect.ShapeCells.LinePattern = 0;


            double max = this.DataPoints.Select(i => i.Value).Max();

            double maxside = 1.5;
            var normalized_values = this.DataPoints.Select(i => i.Value / max).ToList();
            var normalized_widths = normalized_values.Select(i => i*maxside).ToList();

            double cx = margin.Width;

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];
                
                double bar_height = normalized_widths[i];
                double bar_width = normalized_widths[i];

                var bar_ll = new VA.Drawing.Point(cx, lower_y);
                var bar_ur = bar_ll.Add(bar_width, bar_height);

                var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);

                var bar_shape = xdoc.DrawRectangle(bar_rect);
                bar_shape.Text = dp.Value.ToString();
                bar_shape.ShapeCells.LinePattern = 0;
                bar_shape.ShapeCells.LineWeight = 0.0;
                bar_shape.ShapeCells.FillForegnd = ValueColor.ToFormula();
                bar_shape.ShapeCells.VerticalAlign = 0;
                bar_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);

                var label_ll = bar_ll.Subtract(0, margin.Height).Add(0,-0.5);
                var label_ur = label_ll.Add(bar_width, label_height);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);

                var label_shape = xdoc.DrawRectangle(label_rect);
                label_shape.Text = dp.Value.ToString();
                label_shape.ShapeCells.LinePattern = 0;
                label_shape.ShapeCells.LineWeight = 0.0;
                label_shape.ShapeCells.FillPattern= 0;
                label_shape.ShapeCells.VerticalAlign = 0;
                label_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);
                label_shape.Text = dp.Label;

                cx += bar_width + 0.5;


            }

            xdoc.Render(rc.Page);

            return bkrect.Size;

        }
    }

}

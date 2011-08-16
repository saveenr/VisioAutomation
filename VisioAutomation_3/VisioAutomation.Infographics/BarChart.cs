using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public enum BarDirection
    {
        Vertical,
        Horizontal
    }

    public class BarChart : Block
    {
        public IList<DataPoint> DataPoints;
        public VA.Drawing.ColorRGB ValueColor = new VA.Drawing.ColorRGB(0xa0a0a0);
        public VA.Drawing.ColorRGB NonValueColor = new VA.Drawing.ColorRGB(0xffffff);
        double TileHeight = 3.0;
        public BarDirection Direction = BarDirection.Vertical;
        VA.Drawing.Size margin = new VA.Drawing.Size(0.25, 0.25);
        private double _labelHeight = 0.5;
        private double _barDistance = 0.0125;
        private double bar_thickness = 0.5;

        public BarChart()
        {
            this.DataPoints = new List<DataPoint>();
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(rc.PageWidth,TileHeight));

            double lower_y = bkrect.LowerLeft.Y;
            double maxval = 180.0;

            var inner_ll = bkrect.LowerLeft.Add(margin);
            var inner_ur = bkrect.UpperRight.Subtract(margin);
            var innerrect = new VA.Drawing.Rectangle(inner_ll, inner_ur);

            var bararea_ll = innerrect.LowerLeft.Add(0, _labelHeight);
            var bararea_ur = innerrect.UpperRight;
            var bararea_rect = new VA.Drawing.Rectangle(bararea_ll, bararea_ur);
            
            var xdoc = new VA.DOM.Document();

            var tilerect = xdoc.DrawRectangle(bkrect);
            tilerect.ShapeCells.FillForegnd = rc.TileReal.ToFormula();
            tilerect.ShapeCells.LineWeight = 0;
            tilerect.ShapeCells.LinePattern = 0;
            

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];

                var bar_rect = GetBarRect(i, lower_y, maxval, bararea_rect, dp);

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
            if (this.Direction == BarDirection.Vertical)
            {
                var label_ll = bar_rect.LowerLeft.Subtract(0, margin.Height).Add(0, -0.5);
                var label_ur = label_ll.Add(bar_thickness, _labelHeight);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);
                return label_rect;
            }
            else
            {
                double label_width = 2.0;
                var label_ll = bar_rect.LowerLeft.Subtract(margin.Width, 0).Add(-0.5,0).Add(-label_width,0);
                var label_ur = label_ll.Add(label_width, bar_thickness);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);
                return label_rect;
                
            }
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

            if (this.Direction == BarDirection.Vertical)
            {
                bar_shape.ShapeCells.VerticalAlign = 0;
                bar_shape.ShapeCells.HAlign = 0;
            }
            else
            {
                bar_shape.ShapeCells.VerticalAlign = 1;
                bar_shape.ShapeCells.HAlign = 1;
            }
        }

        private Rectangle GetBarRect(int i, double start_pos, double maxval, Rectangle bararea_rect, DataPoint dp)
        {
            if (this.Direction == BarDirection.Vertical)
            {
                double bar_length = dp.Value / maxval * bararea_rect.Height;

                var bar_ll = new VA.Drawing.Point(bar_thickness + _barDistance, start_pos).Multiply(i, 1).Add(margin.Width, margin.Height + _labelHeight);
                var bar_ur = bar_ll.Add(bar_thickness, bar_length);

                var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);
                return bar_rect;
                
            }
            else
            {
                double bar_length = dp.Value / maxval * bararea_rect.Width;

                var bar_ll = new VA.Drawing.Point(start_pos, bar_thickness + _barDistance).Multiply(1, i).Add(margin.Width, margin.Height + _labelHeight);
                var bar_ur = bar_ll.Add(bar_length,bar_thickness);

                var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);
                return bar_rect;
            }
        }
    }

}

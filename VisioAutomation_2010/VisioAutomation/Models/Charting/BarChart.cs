using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    public class BarChart
    {
        public VA.Drawing.Rectangle Rectangle;
        public List<DataPoint> DataPoints;
        public double TotalBarWidth;
        public double TotalMarginWidth;
        public double TotalBarSpacingWidth;
 
        public BarChart(VA.Drawing.Rectangle rect)
        {
            this.Rectangle = rect;
            this.DataPoints = new List<DataPoint>();
        }

        public void Render(IVisio.Page page)
        {
            var values = this.DataPoints.Select(p => p.Value).ToList();

            //var bkshape = page.DrawRectangle(this.Rectangle);

            this.TotalMarginWidth = this.Rectangle.Width*(0.10);
            this.TotalBarSpacingWidth = this.Rectangle.Width * (0.10);
            this.TotalBarWidth = this.Rectangle.Width*(0.80);

            int num_points = this.DataPoints.Count;
            double margin = this.Rectangle.Left + (this.TotalMarginWidth/2.0);
            double bar_spacing = num_points > 1 ? this.TotalBarSpacingWidth/num_points : 0.0;
            double bar_width = num_points > 0 ? this.TotalBarWidth/num_points : this.TotalBarWidth;

            double cur_x = this.Rectangle.Left + margin;

            double max = this.DataPoints.Select(i => i.Value).Max();

            foreach (var dp in this.DataPoints)
            {
                var ll = new VA.Drawing.Point(cur_x, this.Rectangle.Bottom);

                var cur_h = this.Rectangle.Height*(dp.Value/max);
                var up = new VA.Drawing.Point(cur_x + bar_width, this.Rectangle.Bottom + cur_h);

                var bar_rect = new VA.Drawing.Rectangle(ll, up);
                var shape = page.DrawRectangle(bar_rect);
                dp.VisioShape = shape;
                if (dp.Label != null)
                {
                    shape.Text = dp.Label;
                }

                cur_x += bar_width + bar_spacing;
            }

            var allshapes = this.DataPoints.Select(dp => dp.VisioShape).Where(s => s != null).ToList();
            if (allshapes.Count > 0)
            {
                var app = page.Application;
                var win = app.ActiveWindow;
                win.DeselectAll();
                win.DeselectAll();
                win.Select(allshapes, IVisio.VisSelectArgs.visSelect);
                var sel = win.Selection;
                sel.Group();                
            }
        }
    }
}
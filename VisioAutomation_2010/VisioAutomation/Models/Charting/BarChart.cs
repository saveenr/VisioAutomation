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
            this.TotalMarginWidth = this.Rectangle.Width*(0.10);
            this.TotalBarSpacingWidth = this.Rectangle.Width * (0.10);
            this.TotalBarWidth = this.Rectangle.Width*(0.80);

            int num_points = this.DataPoints.Count;
            double margin_pos = this.Rectangle.Left + (this.TotalMarginWidth/2.0);
            double bar_spacing = num_points > 1 ? this.TotalBarSpacingWidth/num_points : 0.0;
            double bar_width = num_points > 0 ? this.TotalBarWidth/num_points : this.TotalBarWidth;

            double cur_x = this.Rectangle.Left + (this.TotalMarginWidth/2.0);

            double max = this.DataPoints.Select(i => i.Value).Max();
            double min = this.DataPoints.Select(i => i.Value).Min();

            double range = -1;
            if (max >= 0)
            {
                if (min >= 0)
                {
                    range = max;
                }
                else
                {
                    range = max - min;
                }
            }
            else
            {
                if (min >= 0)
                {
                    // not possible
                    throw new Exception();
                }
                else
                {
                    range = max - min;
                }
            }


            double base_y = this.Rectangle.Bottom;

            if (min < 0.0)
            {
                base_y += System.Math.Abs(this.Rectangle.Height * (min / range));
            }

            var baseline = page.DrawLine(this.Rectangle.Left, base_y, this.Rectangle.Right, base_y);

            foreach (var dp in this.DataPoints)
            {

                var cur_h = System.Math.Abs(this.Rectangle.Height*(dp.Value/range));

                VA.Drawing.Point ll;
                VA.Drawing.Point up;

                if (dp.Value >= 0.0)
                {
                    ll = new VA.Drawing.Point(cur_x, base_y);
                    up = new VA.Drawing.Point(cur_x + bar_width, base_y + cur_h); ;
                }
                else
                {
                    ll = new VA.Drawing.Point(cur_x, base_y - cur_h);
                    up = new VA.Drawing.Point(cur_x + bar_width, base_y);                    
                }
                
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
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
        public DataPointList DataPoints;
        public double TotalBarWidth;
        public double TotalMarginWidth;
        public double TotalBarSpacingWidth;
 
        public BarChart(VA.Drawing.Rectangle rect)
        {
            this.Rectangle = rect;
            this.DataPoints = new DataPointList();
        }

        public void Render(IVisio.Page page)
        {
            this.TotalMarginWidth = this.Rectangle.Width*(0.10);
            this.TotalBarSpacingWidth = this.Rectangle.Width * (0.10);
            this.TotalBarWidth = this.Rectangle.Width*(0.80);

            int num_points = this.DataPoints.Count;

            double bar_spacing = num_points > 1 ? this.TotalBarSpacingWidth/num_points : 0.0;
            double bar_width = num_points > 0 ? this.TotalBarWidth/num_points : this.TotalBarWidth;

            double cur_x = this.Rectangle.Left + (this.TotalMarginWidth/2.0);

            double max = this.DataPoints.Select(i => i.Value).Max();
            double min = this.DataPoints.Select(i => i.Value).Min();
            var range = ChartUtil.GetValueRangeDistance(min, max);

            double base_y = this.Rectangle.Bottom;

            if (min < 0.0)
            {
                base_y += System.Math.Abs(this.Rectangle.Height * (min / range));
            }

            var category_axis_start_point = new VA.Drawing.Point(this.Rectangle.Left, base_y);
            var category_axis_end_point = new VA.Drawing.Point(this.Rectangle.Right, base_y);
            var category_axis_shape = page.DrawLine(category_axis_start_point, category_axis_end_point);

            foreach (var p in this.DataPoints)
            {
                var value_height = System.Math.Abs(this.Rectangle.Height*(p.Value/range));

                VA.Drawing.Point bar_p0;
                VA.Drawing.Point bar_p1;

                if (p.Value >= 0.0)
                {
                    bar_p0 = new VA.Drawing.Point(cur_x, base_y);
                    bar_p1 = new VA.Drawing.Point(cur_x + bar_width, base_y + value_height); ;
                }
                else
                {
                    bar_p0 = new VA.Drawing.Point(cur_x, base_y - value_height);
                    bar_p1 = new VA.Drawing.Point(cur_x + bar_width, base_y);                    
                }
                
                var bar_rect = new VA.Drawing.Rectangle(bar_p0, bar_p1);
                var shape = page.DrawRectangle(bar_rect);
                p.VisioShape = shape;

                if (p.Label != null)
                {
                    shape.Text = p.Label;
                }

                cur_x += bar_width + bar_spacing;
            }

            var allshapes = this.DataPoints.Select(dp => dp.VisioShape).Where(s => s != null).ToList();
            allshapes.Add(category_axis_shape);

            ChartUtil.GroupShapesIfNeeded(page, allshapes);

        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    public class AreaChart
    {
        public VA.Drawing.Rectangle Rectangle;
        public DataPointList DataPoints;
        public double TotalBarWidth;
        public double TotalMarginWidth;
        public double TotalBarSpacingWidth;
 
        public AreaChart(VA.Drawing.Rectangle rect)
        {
            this.Rectangle = rect;
            this.DataPoints = new DataPointList();
        }

        public void Render(IVisio.Page page)
        {
            this.TotalMarginWidth = this.Rectangle.Width*(0.10);

            int num_points = this.DataPoints.Count;
            double bar_spacing = num_points > 1 ? (this.Rectangle.Width-this.TotalBarWidth)/num_points : 0.0;


            // Calculate min & max which will be used many times later
            double max = this.DataPoints.Select(i => i.Value).Max();
            double min = this.DataPoints.Select(i => i.Value).Min();
            var range = ChartUtil.GetValueRangeDistance(min, max);
            
            // Determine the leftmost part of the drawing area
            double base_x = this.Rectangle.Left + (this.TotalMarginWidth / 2.0);

            // Determine the baseline height a.k.a the y axis location
            double base_y = this.Rectangle.Bottom;
            if (min < 0.0)
            {
                // if the min value is negative then we have to "raise" the baseline to accomodate it
                base_y += System.Math.Abs(this.Rectangle.Height * (min / range));
            }

            var category_axis_start_point = new VA.Drawing.Point(this.Rectangle.Left, base_y);
            var category_axis_end_point = new VA.Drawing.Point(this.Rectangle.Right, base_y);
            var category_axis_shape = page.DrawLine(category_axis_start_point, category_axis_end_point);

            double cur_x = base_x;

            var points = new List<VA.Drawing.Point>();
            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                if (i == 0)
                {
                    points.Add( new VA.Drawing.Point(cur_x,base_y));
                }

                var p = this.DataPoints[i];

                var value_height = System.Math.Abs(this.Rectangle.Height*(p.Value/range));

                if (p.Value >= 0.0)
                {
                    points.Add(new VA.Drawing.Point(cur_x, base_y+value_height));
                }
                else
                {
                    points.Add(new VA.Drawing.Point(cur_x , base_y - value_height));

                }

                if (i == this.DataPoints.Count - 1)
                {
                    points.Add(new VA.Drawing.Point(cur_x, base_y));
                }

                cur_x += bar_spacing;
            }

            points.Add(new VA.Drawing.Point(base_x, base_y));


            var area_shape = page.DrawPolyline(points);
            

            var allshapes = this.DataPoints.Select(dp => dp.VisioShape).Where(s => s != null).ToList();
            allshapes.Add(category_axis_shape);

            ChartUtil.GroupShapesIfNeeded(page, allshapes);
        }

    }
}
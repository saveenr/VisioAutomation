using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public static class DrawingtHelper
    {
        public static IList<IVisio.Shape> DrawPieSlices(IVisio.Page page, VA.Drawing.Point center,
                                                        double radius,
                                                        IList<double> values)
        {
            double sum = values.Sum();
            var shapes = new List<IVisio.Shape>();
            double start_angle = 0;

            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val/sum;
                double cur_angle_size_deg = cur_val_norm*360;
                double end_angle = start_angle + cur_angle_size_deg;
                var pieslice = new VA.Layout.PieSlice(center, radius, start_angle, end_angle);
                var shape = DrawPieSlice(page, pieslice);
                start_angle += cur_angle_size_deg;

                shapes.Add(shape);
            }

            return shapes;
        }

        public static IVisio.Shape DrawPieSlice(
            IVisio.Page page,
            PieSlice pieslice)
        {
            double total_angle = pieslice.EndAngle - pieslice.StartAngle;

            if (total_angle == 0.0)
            {
                var p1 = GetPointAtRadius_Deg(pieslice.Center, pieslice.StartAngle, pieslice.Radius);
                return page.DrawLine(pieslice.Center, p1);
            }
            else if (total_angle >= 360)
            {
                var A = pieslice.Center.Add(-pieslice.Radius, -pieslice.Radius);
                var B = pieslice.Center.Add(pieslice.Radius, pieslice.Radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var shape = page.DrawOval(rect);
                return shape;
            }
            else
            {
                int degree;

                // split apart the arc into distinct bezier segments (will end up with at least 1 segment)
                // the segments will "fit" end to end
                var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                    Convert.DegreesToRadians(pieslice.StartAngle),
                    Convert.DegreesToRadians(pieslice.EndAngle));

                // merge bezier segments together into a list of points
                var merged_points = VA.Drawing.BezierSegment.Merge(sub_arcs, out degree);

                var arc_bez_points = new List<VA.Drawing.Point>(merged_points.Count);
                foreach (var p in merged_points)
                {
                    var np = p.Multiply(pieslice.Radius) + pieslice.Center;
                    arc_bez_points.Add(np);
                }

                // Create one big bezier that accounts for the entire pie shape. This includes the arc
                // calculated above and the sides of the pie slice
                var pie_points = new List<VA.Drawing.Point>(3+arc_bez_points.Count+3);

                var first_point_in_arc = arc_bez_points[0];
                var last_point_in_arc = arc_bez_points[arc_bez_points.Count - 1];

                pie_points.Add(pieslice.Center);
                pie_points.Add(pieslice.Center);
                pie_points.Add(first_point_in_arc);             
                pie_points.AddRange(arc_bez_points);
                pie_points.Add(last_point_in_arc);
                pie_points.Add(pieslice.Center);
                pie_points.Add(pieslice.Center);

                // Render the bezier
                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(pie_points).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        private static VA.Drawing.Point GetPointAtRadius_Deg(VA.Drawing.Point origin, double angle, double radius)
        {
            double theta = VA.Convert.DegreesToRadians(angle);
            double x = radius * System.Math.Cos(theta);
            double y = radius * System.Math.Sin(theta);
            var new_point = new VA.Drawing.Point(x, y);
            new_point = origin + new_point;
            return new_point;
        }

    }
}
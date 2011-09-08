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
                var shape = DrawPieSlice(page, center, radius, start_angle, end_angle);
                start_angle += cur_angle_size_deg;

                shapes.Add(shape);
            }

            return shapes;
        }

        public static IVisio.Shape DrawPieSlice(
            IVisio.Page page, VA.Drawing.Point center, double radius, double start_angle, double end_angle)
        {
            double total_angle = end_angle - start_angle;

            if (total_angle == 0.0)
            {
                var p1 = GetPointAtRadius_Deg(center, start_angle, radius);
                return page.DrawLine(center, p1);
            }
            else if (total_angle >= 360)
            {
                var A = center.Add(-radius, -radius);
                var B = center.Add(radius,   radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var shape = page.DrawOval(rect);
                return shape;
            }
            else
            {
                int degree;
                var arc_bez = GetArcBez(center, radius, start_angle, end_angle, out degree);

                // Create one big bezier that accounts for the entire pie shape. This includes the arc
                // calculated above and the sides of the pie slice
                var pie_bez  = new List<VA.Drawing.Point>(3+arc_bez.Count+3);

                var point_first = arc_bez[0];
                var point_last = arc_bez[arc_bez.Count - 1];

                pie_bez.Add(center);
                pie_bez.Add(center);
                pie_bez.Add(point_first);             
                pie_bez.AddRange(arc_bez);
                pie_bez.Add(point_last);
                pie_bez.Add(center);
                pie_bez.Add(center);

                // Render the bezier
                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(pie_bez).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        private static List<Point> GetArcBez(Point center, double radius, double start_angle, double end_angle, out int degree)
        {
            // split apart the arc into distinct bezier segments (will end up with at least 1 segment)
            // the segments will "fit" end to end
            var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                Convert.DegreesToRadians(start_angle),
                Convert.DegreesToRadians(end_angle));

            // merge bezier segments together into a list of points
            var merged_points = VA.Drawing.BezierSegment.Merge(sub_arcs, out degree);

            var arc_bez = new List<VA.Drawing.Point>(merged_points.Count);
            foreach (var p in merged_points)
            {
                var np = p.Multiply(radius) + center;
                arc_bez.Add(np);
            }
            return arc_bez;
        }

        public static IVisio.Shape DrawArc(
    IVisio.Page page, VA.Drawing.Point center, double inner_radius,  double outer_radius, double start_angle, double end_angle)
        {
            double total_angle = end_angle - start_angle;

            if (total_angle == 0.0)
            {
                var p1 = GetPointAtRadius_Deg(center, start_angle, inner_radius);
                return page.DrawLine(center, p1);
            }
            else if (total_angle >= 360)
            {

                var C = center.Add(-outer_radius, -outer_radius);
                var D = center.Add(outer_radius, outer_radius);
                var outer_rect = new VA.Drawing.Rectangle(C, D);

                var A = center.Add(-inner_radius, -inner_radius).Subtract(C);
                var B = center.Add(inner_radius, inner_radius).Subtract(C);
                var inner_rect = new VA.Drawing.Rectangle(A, B);

                var shape = page.DrawOval(outer_rect);
                shape.DrawOval(inner_rect.Left, inner_rect.Bottom, inner_rect.Right, inner_rect.Top);
                
                return shape;
            }
            else
            {
                int degree;
                var arc_bez_inner = GetArcBez(center, inner_radius, start_angle, end_angle, out degree);
                var arc_bez_outer = GetArcBez(center, outer_radius, start_angle, end_angle, out degree);
                arc_bez_outer.Reverse();

                // Create one big bezier that accounts for the entire pie shape. This includes the arc
                // calculated above and the sides of the pie slice
                var pie_bez = new List<VA.Drawing.Point>(3 + arc_bez_inner.Count + 3);

                var point_first = arc_bez_inner[0];
                var point_last = arc_bez_inner[arc_bez_inner.Count - 1];
                var point_last2 = arc_bez_outer[arc_bez_inner.Count - 1];

                pie_bez.AddRange(arc_bez_inner);

                pie_bez.Add(point_last);
                pie_bez.Add(point_last);

                pie_bez.AddRange(arc_bez_outer);

                pie_bez.Add(point_last2);
                pie_bez.Add(point_first);
                pie_bez.Add(point_first);

                // Render the bezier
                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(pie_bez).ToArray();
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
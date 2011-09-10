using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation
{
    public struct Angle
    {
        private double theta_rad;

        private Angle(double v)
        {
            this.theta_rad = v;
        }

        public static Angle FromRadians(double v)
        {
            return new Angle(v);
        }

        public double Radians
        {
            get { return this.theta_rad; }
        }

        public double Degrees
        {
            get { return VA.Convert.RadiansToDegrees(this.theta_rad); }
        }

        public static Angle operator +(Angle x, Angle y)
        {
            return new VA.Angle(x.Radians + y.Radians);
        }

        public static Angle operator -(Angle x, Angle y)
        {
            return new VA.Angle(x.Radians - y.Radians);
        }

        public static implicit operator Angle (double r)
        {
            return new Angle(r);
        }
    }
}


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
            Angle start_angle = 0;

            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val/sum;
                Angle cur_angle = cur_val_norm * System.Math.PI * 2.0;
                Angle end_angle = start_angle.Radians + cur_angle.Radians;
                var shape = DrawPieSlice(page, center, radius, start_angle, end_angle);
                start_angle += cur_angle;

                shapes.Add(shape);
            }

            return shapes;
        }

        public static IVisio.Shape DrawPieSlice(
            IVisio.Page page, VA.Drawing.Point center, double radius, VA.Angle start_angle, VA.Angle end_angle)
        {
            double total_angle = end_angle.Radians - start_angle.Radians;

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
                var pie_bez = GetPieSliceBezier(center, radius, start_angle, end_angle, out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(pie_bez).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        private static List<Point> GetPieSliceBezier(Point center, double radius, VA.Angle start_angle, VA.Angle end_angle, out int degree)
        {
            var arc_bez = GetArcBez(center, radius, start_angle, end_angle, out degree);

            // Create one big bezier that accounts for the entire pie shape. This includes the arc
            // calculated above and the sides of the pie slice
            var pie_bez = new List<VA.Drawing.Point>(3 + arc_bez.Count + 3);

            var point_first = arc_bez[0];
            var point_last = arc_bez[arc_bez.Count - 1];

            pie_bez.Add(center);
            pie_bez.Add(center);
            pie_bez.Add(point_first);
            pie_bez.AddRange(arc_bez);
            pie_bez.Add(point_last);
            pie_bez.Add(center);
            pie_bez.Add(center);
            return pie_bez;
        }

        private static List<Point> GetArcBez(Point center, double radius, VA.Angle start_angle, VA.Angle end_angle, out int degree)
        {
            // split apart the arc into distinct bezier segments (will end up with at least 1 segment)
            // the segments will "fit" end to end
            var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                start_angle,
                end_angle);

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
            IVisio.Page page, VA.Drawing.Point center, double inner_radius, double outer_radius, VA.Angle start_angle, VA.Angle end_angle)
        {
            VA.Angle total_angle = end_angle - start_angle;

            if (total_angle.Radians == 0.0)
            {
                var p1 = GetPointAtRadius_Deg(center, start_angle, inner_radius);
                var p2 = GetPointAtRadius_Deg(center, start_angle, outer_radius);
                var shape = page.DrawLine(p1, p2);
                return shape;
            }
            else if (total_angle.Radians >= 360)
            {
                var outer_radius_point = new VA.Drawing.Point(outer_radius, outer_radius);
                var C = center - outer_radius_point;
                var D = center + outer_radius_point;
                var outer_rect = new VA.Drawing.Rectangle(C, D);

                var inner_radius_point = new VA.Drawing.Point(inner_radius, inner_radius);
                var A = center - inner_radius_point - C;
                var B = center + inner_radius_point - C;
                var inner_rect = new VA.Drawing.Rectangle(A, B);

                var shape = page.DrawOval(outer_rect);
                shape.DrawOval(inner_rect.Left, inner_rect.Bottom, inner_rect.Right, inner_rect.Top);
                
                return shape;
            }
            else
            {
                int degree;
                var thickarc = GetThinkArcBezier(center, inner_radius, outer_radius, start_angle, end_angle, out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.DrawingUtil.PointsToDoubles(thickarc).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        private static List<Point> GetThinkArcBezier(Point center, double inner_radius, double outer_radius, VA.Angle start_angle, VA.Angle end_angle, out int degree)
        {
            var bez_inner = GetArcBez(center, inner_radius, start_angle, end_angle, out degree);
            var bez_outer = GetArcBez(center, outer_radius, start_angle, end_angle, out degree);
            bez_outer.Reverse();

            // Create one big bezier that accounts for the entire pie shape. This includes the arc
            // calculated above and the sides of the pie slice
            var bez = new List<VA.Drawing.Point>(3 + bez_inner.Count + 3);

            var point_first = bez_inner[0];
            var point_last = bez_inner[bez_inner.Count - 1];
            var point_last2 = bez_outer[bez_inner.Count - 1];

            bez.AddRange(bez_inner);

            bez.Add(point_last);
            bez.Add(point_last);

            bez.AddRange(bez_outer);

            bez.Add(point_last2);
            bez.Add(point_first);
            bez.Add(point_first);
            return bez;
        }


        private static VA.Drawing.Point GetPointAtRadius_Deg(VA.Drawing.Point origin, VA.Angle angle, double radius)
        {
            double theta = angle.Radians;
            double x = radius * System.Math.Cos(theta);
            double y = radius * System.Math.Sin(theta);
            var new_point = new VA.Drawing.Point(x, y);
            new_point = origin + new_point;
            return new_point;
        }

    }
}
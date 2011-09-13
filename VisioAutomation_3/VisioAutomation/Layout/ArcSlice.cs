using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public class ArcSlice
    {
        public VA.Drawing.Point Center;
        public double StartAngle;
        public double EndAndle;
        public double InnerRadius;
        public double OuterRadius;

        public ArcSlice(VA.Drawing.Point center, double innerRadius, double outerRadius, double start, double end)
        {
            this.Center = center;
            this.InnerRadius = innerRadius;
            this.OuterRadius = outerRadius;
            this.StartAngle = start;
            this.EndAndle = end;
        }

        public double Angle
        {
            get { return this.EndAndle - this.StartAngle; }
        }

        public IVisio.Shape Render(IVisio.Page page)
        {
            double total_angle = this.Angle;

            if (total_angle == 0.0)
            {
                var p1 = this.GetPointAtRadius(this.Center, this.StartAngle, this.InnerRadius);
                var p2 = this.GetPointAtRadius(this.Center, this.StartAngle, this.OuterRadius);
                var shape = page.DrawLine(p1, p2);
                return shape;
            }
            else if (total_angle >= System.Math.PI)
            {
                var outer_radius_point = new VA.Drawing.Point(this.OuterRadius, this.OuterRadius);
                var C = this.Center - outer_radius_point;
                var D = this.Center + outer_radius_point;
                var outer_rect = new VA.Drawing.Rectangle(C, D);

                var inner_radius_point = new VA.Drawing.Point(this.InnerRadius, this.InnerRadius);
                var A = this.Center - inner_radius_point - C;
                var B = this.Center + inner_radius_point - C;
                var inner_rect = new VA.Drawing.Rectangle(A, B);

                var shape = page.DrawOval(outer_rect);
                shape.DrawOval(inner_rect.Left, inner_rect.Bottom, inner_rect.Right, inner_rect.Top);

                return shape;
            }
            else
            {
                int degree;
                var thickarc = GetThinArcBezier(this.Center, this.InnerRadius, this.OuterRadius, this.StartAngle, this.EndAndle, out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.Point.ToDoubles(thickarc).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        static List<Point> GetThinArcBezier(Point center, double inner_radius, double outer_radius, double start_angle, double end_angle, out int degree)
        {
            var bez_inner = VA.Layout.ArcSlice.GetArcBez(center, inner_radius, start_angle, end_angle, out degree);
            var bez_outer = VA.Layout.ArcSlice.GetArcBez(center, outer_radius, start_angle, end_angle, out degree);
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

        internal static List<Point> GetArcBez(Point center, double radius, double start_angle, double end_angle, out int degree)
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

        VA.Drawing.Point GetPointAtRadius(VA.Drawing.Point origin, double angle, double radius)
        {
            double x = radius * System.Math.Cos(angle);
            double y = radius * System.Math.Sin(angle);
            var new_point = new VA.Drawing.Point(x, y);
            new_point = origin + new_point;
            return new_point;
        }

    }
}
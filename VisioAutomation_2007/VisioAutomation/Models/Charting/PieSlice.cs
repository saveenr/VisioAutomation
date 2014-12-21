using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    public class PieSlice
    {
        public double InnerRadius { get; private set; }
        public double Radius { get; private set; }
        public VA.Drawing.Point Center { get; private set; }
        public double SectorStartAngle { get; private set; }
        public double SectorEndAngle { get; private set; }

        public double Angle
        {
            get { return this.SectorEndAngle - this.SectorStartAngle; }
        }

        public PieSlice(VA.Drawing.Point center, double start, double end)
        {
            this.Center = center;

            if (end < start)
            {
                throw new System.ArgumentException("end","end angle must be greater than or equal to start angle");
            }

            this.SectorStartAngle = start;
            this.SectorEndAngle = end;
        }

        public PieSlice(VA.Drawing.Point center, double radius, double start, double end) :
            this(center,start,end)
        {
            if (radius < 0.0)
            {
                throw new System.ArgumentOutOfRangeException("radius","must be non-negative");
            }

            this.Radius = radius;
        }

        public PieSlice(VA.Drawing.Point center, double start, double end, double inner_radius, double radius) :
            this(center,start,end)
        {
            if (inner_radius < 0.0)
            {
                throw new System.ArgumentException("inner_radius", "must be non-negative");
            }

            if (radius < 0.0)
            {
                throw new System.ArgumentException("outer_radius", "must be non-negative");
            }

            if (inner_radius > radius)
            {
                throw new System.ArgumentException("inner_radius", "must be less than or equal to outer_radius");                
            }

            this.InnerRadius = inner_radius;
            this.Radius = radius;
        }


        internal List<VA.Drawing.Point> GetShapeBezierForPie(out int degree)
        {
            this.check_normal_angle();

            var arc_bez = this.GetArcBez(this.Radius, out degree);

            // Create one big bezier that accounts for the entire pie shape. This includes the arc
            // calculated above and the sides of the pie slice
            var pie_bez = new List<VA.Drawing.Point>(3 + arc_bez.Count + 3);

            var point_first = arc_bez[0];
            var point_last = arc_bez[arc_bez.Count - 1];

            pie_bez.Add(this.Center);
            pie_bez.Add(this.Center);
            pie_bez.Add(point_first);
            pie_bez.AddRange(arc_bez);
            pie_bez.Add(point_last);
            pie_bez.Add(this.Center);
            pie_bez.Add(this.Center);
            return pie_bez;
        }

        public IVisio.Shape Render(IVisio.Page page)
        {
            if (InnerRadius <= 0.0)
            {
                return this.RenderPie(page);
            }
            else
            {
                return this.RenderDoughnut(page);
            }
        }

        public IVisio.Shape RenderPie( IVisio.Page page)
        {
            if (this.Angle == 0.0)
            {
                var p1 = this.GetPointAtRadius(this.Center, this.Radius, this.SectorStartAngle);
                return page.DrawLine(this.Center, p1);
            }
            else if (this.Angle >= 2*System.Math.PI)
            {
                var A = this.Center.Add(-this.Radius, -this.Radius);
                var B = this.Center.Add(this.Radius, this.Radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var shape = page.DrawOval(rect);
                return shape;
            }
            else
            {
                int degree;
                var pie_bez = this.GetShapeBezierForPie(out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.Point.ToDoubles(pie_bez).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        public IVisio.Shape RenderDoughnut(IVisio.Page page)
        {
            double total_angle = this.Angle;

            if (total_angle == 0.0)
            {
                var p1 = this.GetPointAtRadius(this.Center, this.SectorStartAngle, this.InnerRadius);
                var p2 = this.GetPointAtRadius(this.Center, this.SectorStartAngle, this.Radius);
                var shape = page.DrawLine(p1, p2);
                return shape;
            }
            else if (total_angle >= System.Math.PI)
            {
                var outer_radius_point = new VA.Drawing.Point(this.Radius, this.Radius);
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
                var thickarc = this.GetShapeBezierForDoughnut(out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.Point.ToDoubles(thickarc).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        List<VA.Drawing.Point> GetShapeBezierForDoughnut(out int degree)
        {
            this.check_normal_angle();

            var bez_inner = this.GetArcBez(this.InnerRadius, out degree);
            var bez_outer = this.GetArcBez(this.Radius, out degree);
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

        public static List<PieSlice> GetSlicesFromValues(VA.Drawing.Point center, double radius, IList<double> values)
        {
            double sectors_sum = values.Sum();
            var slices = new List<PieSlice>(values.Count);
            double start_angle = 0;
            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val / sectors_sum;
                double cur_angle = cur_val_norm * System.Math.PI * 2.0;
                double end_angle = start_angle + cur_angle;

                var ps = new VA.Models.Charting.PieSlice(center,radius,start_angle, end_angle);
                slices.Add(ps);

                start_angle += cur_angle;
            }
            return slices;
        }

        public static List<PieSlice> GetSlicesFromValues(VA.Drawing.Point center, double inner_radius, double outer_radius, IList<double> values)
        {
            var slices = GetSlicesFromValues(center, outer_radius, values);
            foreach (var slice in slices)
            {
                slice.InnerRadius = inner_radius;
            }
            return slices;
        }

        protected VA.Drawing.Point GetPointAtRadius(VA.Drawing.Point origin, double angle, double radius)
        {
            double x = radius * System.Math.Cos(angle);
            double y = radius * System.Math.Sin(angle);
            var new_point = new VA.Drawing.Point(x, y);
            new_point = origin + new_point;
            return new_point;
        }

        protected List<VA.Drawing.Point> GetArcBez(double radius, out int degree)
        {
            // split apart the arc into distinct bezier segments (will end up with at least 1 segment)
            // the segments will "fit" end to end
            var sub_arcs = VA.Drawing.BezierSegment.FromArc(
                this.SectorStartAngle,
                this.SectorEndAngle);

            // merge bezier segments together into a list of points
            var merged_points = VA.Drawing.BezierSegment.Merge(sub_arcs, out degree);

            var arc_bez = new List<VA.Drawing.Point>(merged_points.Count);
            foreach (var p in merged_points)
            {
                var np = p.Multiply(radius) + this.Center;
                arc_bez.Add(np);
            }
            return arc_bez;
        }

        protected void check_normal_angle()
        {
            if ((this.Angle <= 0.0) || (this.Angle > System.Math.PI * 2.0))
            {
                string msg = string.Format("Angle of sector must be greater than zero and less than 2*PI");
                throw new System.ArgumentException(msg);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{

    public class RadialSlice
    {
        public VA.Drawing.Point Center { get; private set; }
        public double StartAngle { get; private set; }
        public double EndAndle { get; private set; }

        public RadialSlice(VA.Drawing.Point center, double start, double end)
        {
            this.Center = center;
            this.StartAngle = start;
            this.EndAndle = end;            
        }

        public double Angle
        {
            get { return this.EndAndle - this.StartAngle; }
        }

        protected VA.Drawing.Point GetPointAtRadius(VA.Drawing.Point origin, double angle, double radius)
        {
            double x = radius * System.Math.Cos(angle);
            double y = radius * System.Math.Sin(angle);
            var new_point = new VA.Drawing.Point(x, y);
            new_point = origin + new_point;
            return new_point;
        }

        protected static List<Point> GetArcBez(Point center, double radius, double start_angle, double end_angle, out int degree)
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


        protected static List<RadialSlice> GetSlicesFromValues(Point center, IList<double> values)
        {
            double sum = values.Sum();
            var slices = new List<VA.Layout.RadialSlice>(values.Count);
            double start_angle = 0;
            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val / sum;
                double cur_angle = cur_val_norm * System.Math.PI * 2.0;
                double end_angle = start_angle + cur_angle;

                var ps = new VA.Layout.RadialSlice(center, start_angle, end_angle);
                slices.Add(ps);

                start_angle += cur_angle;
            }
            return slices;
        }

    }
}
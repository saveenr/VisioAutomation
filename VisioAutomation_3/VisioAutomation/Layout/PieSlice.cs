using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{
    public class PieSlice : RadialSlice
    {
        public double Radius { get; private set; }

        public PieSlice(Point center, double start, double end, double radius) :
            base(center,start,end)
        {
            this.Radius = radius;
        }

        internal static List<Point> GetPieSliceBezier(Point center, double radius, double start_angle, double end_angle, out int degree)
        {
            var arc_bez = RadialSlice.GetArcBez(center, radius, start_angle, end_angle, out degree);

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

        public IVisio.Shape Render( IVisio.Page page)
        {

            if (this.Angle == 0.0)
            {
                var p1 = this.GetPointAtRadius(this.Center, this.StartAngle, this.Radius);
                return page.DrawLine(this.Center, p1);
            }
            else if (this.Angle >= System.Math.PI)
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
                var pie_bez = GetPieSliceBezier(this.Center, this.Radius, this.StartAngle, this.EndAndle, out degree);

                // Render the bezier
                var doubles_array = VA.Drawing.Point.ToDoubles(pie_bez).ToArray();
                var pie_slice = page.DrawBezier(doubles_array, (short)degree, 0);
                return pie_slice;
            }
        }

        public static IList<IVisio.Shape> DrawPieSlices(IVisio.Page page, VA.Drawing.Point center, double radius, IList<double> values)
        {
            var slices = GetSlicesFromValues(center, radius, values);
            var shapes = new List<IVisio.Shape>(slices.Count);

            foreach (var slice in slices)
            {
                var shape = slice.Render(page);
                shapes.Add(shape);
            }

            return shapes;
        }

        public static List<PieSlice> GetSlicesFromValues(Point center, double radius, IList<double> values)
        {
            var base_slices = RadialSlice.GetSlicesFromValues(center, values);
            var slices = new List<VA.Layout.PieSlice>(values.Count);
            foreach (var base_slice in base_slices)
            {
                var slice = new PieSlice(base_slice.Center, base_slice.StartAngle, base_slice.EndAndle, radius);
                slices.Add(slice);
            }
            return slices;
        }
    }
}
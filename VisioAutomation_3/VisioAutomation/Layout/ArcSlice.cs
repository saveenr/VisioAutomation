using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout
{

    public class ArcSlice : RadialSlice
    {
        public double InnerRadius { get; private set; }
        public double OuterRadius { get; private set; }

        public ArcSlice(Point center, double start, double end, double innerRadius, double outerRadius) :
            base(center,start,end)
        {
            this.InnerRadius = innerRadius;
            this.OuterRadius = outerRadius;
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

    }
}
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Models.Radial
{
    public class RadialSlice
    {
        public VA.Drawing.Point Center { get; private set; }
        public VA.Layout.Models.Radial.Sector Sector { get; private set; }

        public RadialSlice(VA.Drawing.Point center, double start, double end)
        {
            this.Center = center;

            if (end < start)
            {
                throw new System.ArgumentException("end","end angle must be greater than or equal to start angle");
            }

            this.Sector = new VA.Layout.Models.Radial.Sector(start, end);
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
                this.Sector.StartAngle,
                this.Sector.EndAngle);

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

        protected static List<T> GetSlicesFromValues<T>(VA.Drawing.Point center, IList<double> values, System.Func<Sector, T> make_slice)
        {
            var sectors = RadialSlice.GetSectorsFromValues(values);
            var slices = new List<T>(values.Count);
            foreach (var sector in sectors)
            {
                var s = make_slice(sector);
                slices.Add(s);
            }
            return slices;
        }

        private static List<Sector> GetSectorsFromValues(IList<double> values)
        {
            double sectors = values.Sum();
            var slices = new List<Sector>(values.Count);
            double start_angle = 0;
            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val / sectors;
                double cur_angle = cur_val_norm * System.Math.PI * 2.0;
                double end_angle = start_angle + cur_angle;

                var ps = new VA.Layout.Models.Radial.Sector(start_angle, end_angle);
                slices.Add(ps);

                start_angle += cur_angle;
            }
            return slices;
        }

        protected void check_normal_angle()
        {
            if ((this.Sector.Angle <= 0.0) || (this.Sector.Angle > System.Math.PI * 2.0))
            {
                string msg = string.Format("Angle of sector must be greater than zero and less than 2*PI");
                throw new System.ArgumentException(msg);
            }
        }
    }
}
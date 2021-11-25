﻿using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Models.Geometry
{
    public struct BezierSegment
    {
        public VisioAutomation.Geometry.Point Start { get; }
        public VisioAutomation.Geometry.Point Handle1 { get; }
        public VisioAutomation.Geometry.Point Handle2 { get; }
        public VisioAutomation.Geometry.Point End { get; }

        public BezierSegment(VisioAutomation.Geometry.Point start, VisioAutomation.Geometry.Point handle1, VisioAutomation.Geometry.Point handle2, VisioAutomation.Geometry.Point end)
            : this()
        {
            this.Start = start;
            this.Handle1 = handle1;
            this.Handle2 = handle2;
            this.End = end;
        }

        public BezierSegment(IList<VisioAutomation.Geometry.Point> points)
            : this()
        {
            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            if (points.Count != 4)
            {
                string msg = string.Format("A {0} must have exactly 4 points", nameof(BezierSegment));
                throw new System.ArgumentException(msg, nameof(points));
            }

            this.Start = points[0];
            this.Handle1 = points[1];
            this.Handle2 = points[2];
            this.End = points[3];
        }

        public static List<VisioAutomation.Geometry.Point> Merge(IList<BezierSegment> segments, out int degree)
        {
            if (segments == null)
            {
                throw new System.ArgumentNullException(nameof(segments));
            }

            var points = new List<VisioAutomation.Geometry.Point>(segments.Count * 4);
            int n = 0;
            foreach (var seg in segments)
            {
                if (n == 0)
                {
                    points.Add(seg.Start);
                }

                points.Add(seg.Handle1);
                points.Add(seg.Handle2);
                points.Add(seg.End);
                n++;
            }

            degree = 3;
            return points;
        }

        public static BezierSegment[] FromArc(double startangle, double endangle)
        {
            if (endangle < startangle)
            {
                throw new System.ArgumentOutOfRangeException(nameof(endangle), "endangle must be >= startangle");
            }

            double min_angle = 0;
            double max_angle = System.Math.PI * 2;
            double  total_angle = endangle - startangle;

            if (total_angle == min_angle)
            {
                var arr = new BezierSegment[1];
                double cos_theta = System.Math.Cos(startangle);
                double sin_theta = System.Math.Sin(startangle);
                var p0 = new VisioAutomation.Geometry.Point(cos_theta, -sin_theta);
                var p1 = BezierSegment._rotate_around_origin( p0, startangle);
                arr[0] = new BezierSegment(p1,p1,p1,p1);
            }

            if (total_angle > max_angle)
            {
                endangle = startangle + max_angle;
            }

            var bez_arr = BezierSegment.subdivide_arc_nicely(startangle, endangle)
                .Select(a => BezierSegment.get_bezier_points_for_small_arc(a.Begin, a.End))
                .ToArray();

            return bez_arr;
        }

        private static IEnumerable<ArcSegment> subdivide_arc_nicely(double start_angle, double end_angle)
        {
            // TODO: Should calculate number of subarcs without resorting to an enumeration

            if (start_angle > end_angle)
            {
                throw new System.ArgumentException("end_angle must be < than start angle", nameof(end_angle));
            }

            // the original purpose of this method is to break apart arcs > 90 degrees into smaller sub-arcs of 90 or less
            // the current implementation does that but also fractures the arc on 0,90,180,270, etc. degrees even if the arc length is <90
            // example (85,110) -> (85,90) & (90,110)

            double right_angle = System.Math.PI/2;

            var cur = new ArcSegment(start_angle, end_angle);
            while (true)
            {
                double temp = System.Math.Floor(cur.Begin/right_angle);
                double cut_angle = (temp + 1)*right_angle;

                if ((cur.Begin < cut_angle) && (cut_angle < cur.End))
                {
                    yield return (new ArcSegment(cur.Begin, cut_angle));
                    cur = new ArcSegment(cut_angle, cur.End);
                }
                else
                {
                    yield return cur;
                    break;
                }
            }
        }

        private static BezierSegment get_bezier_points_for_small_arc(double start_angle, double end_angle)
        {
            const double right_angle = System.Math.PI/2;
            double total_angle = end_angle - start_angle;

            if (total_angle > right_angle)
            {
                throw new System.ArgumentOutOfRangeException(nameof(end_angle),
                                                             "angle formed by start and end must <= right angle (pi/2)");
            }

            double theta = (end_angle - start_angle)/2;
            double cos_theta = System.Math.Cos(theta);
            double sin_theta = System.Math.Sin(theta);

            var p0 = new VisioAutomation.Geometry.Point(cos_theta, -sin_theta);
            var p1 = new VisioAutomation.Geometry.Point((4 - cos_theta) / 3.0, ((1 - cos_theta) * (cos_theta - 3.0)) / (3 * sin_theta));
            var p2 = new VisioAutomation.Geometry.Point(p1.X, -p1.Y);
            var p3 = new VisioAutomation.Geometry.Point(p0.X, -p0.Y);

            var arc_bezier = new[] {p0, p1, p2, p3}
                .Select(p => BezierSegment._rotate_around_origin(p, theta + start_angle))
                .ToArray();

            return new BezierSegment(arc_bezier);
        }

        private static VisioAutomation.Geometry.Point _rotate_around_origin(VisioAutomation.Geometry.Point p1, double theta)
        {
            double nx = (System.Math.Cos(theta)*p1.X) - (System.Math.Sin(theta)*p1.Y);
            double ny = (System.Math.Sin(theta)*p1.X) + (System.Math.Cos(theta)*p1.Y);
            return new VisioAutomation.Geometry.Point(nx, ny);
        }
    }
}
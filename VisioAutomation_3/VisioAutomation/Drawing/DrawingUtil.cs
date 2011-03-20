using System;
using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public static class DrawingUtil
    {
        public static VA.Drawing.Point Max(VA.Drawing.Point a, VA.Drawing.Point b)
        {
            return new Drawing.Point(Math.Max(a.X, b.X),
                             Math.Max(a.Y, b.Y));
        }

        public static VA.Drawing.Point Min(VA.Drawing.Point a, VA.Drawing.Point b)
        {
            return new VA.Drawing.Point(Math.Min(a.X, b.X),
                             Math.Min(a.Y, b.Y));
        }

        public static VA.Drawing.Size Max(VA.Drawing.Size a, VA.Drawing.Size b)
        {
            return new VA.Drawing.Size(Math.Max(a.Width, b.Width),
                            Math.Max(a.Height, b.Height));
        }

        public static VA.Drawing.Size Min(VA.Drawing.Size a, VA.Drawing.Size b)
        {
            return new VA.Drawing.Size(Math.Min(a.Width, b.Width),
                            Math.Min(a.Height, b.Height));
        }

        public static VA.Drawing.Size SnapToNearestValue(VA.Drawing.Size size, VA.Drawing.Size snapsize)
        {
            return new VA.Drawing.Size(VA.Internal.MathUtil.Round(size.Width, snapsize.Width),
                            VA.Internal.MathUtil.Round(size.Height, snapsize.Height));
        }

        public static IEnumerable<Drawing.Point> DoublesToPoints(IEnumerable<double> doubles)
        {
            if (doubles == null)
            {
                throw new ArgumentNullException("doubles");
            }

            int count = 0;
            double even_value = default(double);
            foreach (var value in doubles)
            {
                if ((count%2) == 0)
                {
                    even_value = value;
                }
                else
                {
                    yield return new Drawing.Point(even_value, value);
                }
                count++;
            }
        }

        public static IEnumerable<double> PointsToDoubles(IEnumerable<Drawing.Point> points)
        {
            foreach (var p in points)
            {
                yield return p.X;
                yield return p.Y;
            }
        }

        public static VA.Drawing.Point Round(VA.Drawing.Point p, double xd, double yd)
        {
            return new Drawing.Point(VA.Internal.MathUtil.Round(p.X, xd),
                             VA.Internal.MathUtil.Round(p.Y, yd));
        }

        public static VA.Drawing.Rectangle GetBoundingBox(VA.Drawing.Rectangle r1, VA.Drawing.Rectangle r2)
        {
            double left = Math.Min(r1.Left, r2.Left);
            double bottom = Math.Min(r1.Bottom, r2.Bottom);
            double right = Math.Max(r1.Right, r2.Right);
            double top = Math.Max(r1.Top, r2.Top);

            var r = new VA.Drawing.Rectangle(left, bottom, right, top);
            return r;
        }

        public static VA.Drawing.Rectangle GetBoundingBox(Drawing.Point p1, Drawing.Point p2)
        {
            double left = Math.Min(p1.X, p2.X);
            double bottom = Math.Min(p1.Y, p2.Y);
            double right = Math.Max(p1.X, p2.X);
            double top = Math.Max(p1.X, p2.X);

            var r = new Drawing.Rectangle(left, bottom, right, top);
            return r;
        }

        public static VA.Drawing.Rectangle? TryGetBoundingBox(IEnumerable<Drawing.Rectangle> rects)
        {
            int count = 0;
            VA.Drawing.Rectangle old_val = default(VA.Drawing.Rectangle);

            int rect_index = 0;
            foreach (var rect in rects)
            {
                old_val = rect_index == 0 ? rect : GetBoundingBox(old_val, rect);
                count++;
                rect_index++;
            }
            if (count > 0)
            {
                return old_val;
            }
            return null;
        }

        public static VA.Drawing.Rectangle GetBoundingBox(IEnumerable<Drawing.Point> points)
        {
            var result = TryGetBoundingBox(points);
            if (!result.HasValue)
            {
                string msg = "Failed to create bounding box from points";
                throw new System.ArgumentException(msg, "points");
            }
            return result.Value;
        }

        public static VA.Drawing.Rectangle GetBoundingBox(IEnumerable<VA.Drawing.Rectangle> rects)
        {
            var result = TryGetBoundingBox(rects);
            if (!result.HasValue)
            {
                string msg = "Failed to create bounding box from rects";
                throw new System.ArgumentException(msg, "rects");
            }
            else
            {
                return result.Value;
            }
        }

        public static VA.Drawing.Rectangle? TryGetBoundingBox(IEnumerable<VA.Drawing.Point> points)
        {
            int count = 0;
            VA.Drawing.Point min = default(Drawing.Point);
            VA.Drawing.Point max = default(Drawing.Point);
            int point_index = 0;
            foreach (var point in points)
            {
                if (point_index == 0)
                {
                    min = point;
                    max = point;
                }
                else
                {
                    min = Min(min, point);
                    max = Max(max, point);
                }
                count++;
                point_index++;
            }
            if (count > 0)
            {
                return new VA.Drawing.Rectangle(min.X, min.Y, max.X, max.Y);
            }
            return null;
        }
    }
}
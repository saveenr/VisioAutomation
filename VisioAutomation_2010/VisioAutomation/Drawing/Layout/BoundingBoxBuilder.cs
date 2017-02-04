using System.Collections.Generic;

namespace VisioAutomation.Drawing.Layout
{
    public static class BoundingBoxBuilder
    {
        public static Rectangle? FromPoints(IEnumerable<Point> points)
        {
            bool initialized = false;

            double _min_x = 0.0;
            double _min_y = 0.0;
            double _max_x = 0.0;
            double _max_y = 0.0;

            foreach (var p in points)
            {
                if (initialized)
                {
                    _min_x = System.Math.Min(_min_x, p.X);
                    _max_x = System.Math.Max(_max_x, p.X);
                    _min_y = System.Math.Min(_min_y, p.Y);
                    _max_y = System.Math.Max(_max_y, p.Y);
                }
                else
                {
                    _min_x = p.X;
                    _max_x = p.X;
                    _min_y = p.Y;
                    _max_y = p.Y;
                    initialized = true;
                }
            }

            if (initialized)
            {
                return new Rectangle(_min_x, _min_y, _max_x, _max_y);
            }

            return null;
        }

        private static IEnumerable<Point> rects_to_points(IEnumerable<Rectangle> rects)
        {
            foreach (var r in rects)
            {
                yield return r.LowerLeft;
                yield return r.UpperRight;
            }
        }

        public static Rectangle? FromRectangles(IEnumerable<Rectangle> rects)
        {
            var points = rects_to_points(rects);
            return BoundingBoxBuilder.FromPoints(points);
        }
    }
}
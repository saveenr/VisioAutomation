using System.Collections.Generic;

namespace VisioAutomation.Models.Geometry
{
    public class BoundingBoxBuilder
    {
        bool _initialized = false;
        double _min_x = 0.0;
        double _min_y = 0.0;
        double _max_x = 0.0;
        double _max_y = 0.0;

        public BoundingBoxBuilder()
        {
            
        }

        public void Add(VisioAutomation.Core.Point p)
        {

            if (_initialized)
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
                _initialized = true;
            }
        }

        public void Add(VisioAutomation.Core.Rectangle r)
        {
            this.Add(r.LowerLeft);
            this.Add(r.UpperRight);
        }

        public VisioAutomation.Core.Rectangle? ToRectangle()
        {
            if (_initialized)
            {
                return new VisioAutomation.Core.Rectangle(_min_x, _min_y, _max_x, _max_y);
            }

            return null;
        }

        public void AddRange(IEnumerable<VisioAutomation.Core.Point> points)
        {
            foreach (var p in points)
            {
                this.Add(p);
            }
        }

        public void AddRange(IEnumerable<VisioAutomation.Core.Rectangle> rects)
        {
            foreach (var r in rects)
            {
                this.Add(r);
            }
        }

        public static VisioAutomation.Core.Rectangle? FromRectangles(IEnumerable<VisioAutomation.Core.Rectangle> rects)
        {
            var bbb = new BoundingBoxBuilder();
            bbb.AddRange(rects);
            return bbb.ToRectangle();
        }

        public static VisioAutomation.Core.Rectangle? FromPoints(IEnumerable<VisioAutomation.Core.Point> points)
        {
            var bbb = new BoundingBoxBuilder();
            bbb.AddRange(points);
            return bbb.ToRectangle();
        }

    }
}
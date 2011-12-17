using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct LineSegment
    {
        private readonly VA.Drawing.Point _start;
        private readonly VA.Drawing.Point _end;

        public LineSegment(VA.Drawing.Point start, VA.Drawing.Point end)
        {
            this._start = start;
            this._end = end;
        }

        public LineSegment(VA.Drawing.Point[] points)
        {
            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            if (points.Length != 2)
            {
                throw new System.ArgumentOutOfRangeException("points", "Must have exactly 2 points");
            }

            this._start = points[0];
            this._end = points[1];
        }

        public VA.Drawing.Point Start
        {
            get { return _start; }
        }

        public VA.Drawing.Point End
        {
            get { return _end; }
        }

        public VA.Drawing.Point[] ToPoints()
        {
            var points = new[] { this._start, this._end };
            return points;
        }
    }
}
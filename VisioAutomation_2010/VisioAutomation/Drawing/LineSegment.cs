using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public struct LineSegment
    {
        private readonly Point _start;
        private readonly Point _end;

        public LineSegment(Point start, Point end)
        {
            this._start = start;
            this._end = end;
        }

        public Point Start
        {
            get { return this._start; }
        }

        public Point End
        {
            get { return this._end; }
        }
    }
}
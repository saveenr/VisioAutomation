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

        public VA.Drawing.Point Start
        {
            get { return _start; }
        }

        public VA.Drawing.Point End
        {
            get { return _end; }
        }
    }
}
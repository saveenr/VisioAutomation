using System;
using VisioAutomation.Utilities;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class FixedSrcStreamBuilder : FixedStreamBuilder<Src>
    {
        public FixedSrcStreamBuilder(int capacity) : base(capacity,3)
        {

        }

        protected override void _Add(Utilities.ArraySegment<short> seg, Src item)
        {
            seg[0] = item.Section;
            seg[1] = item.Row;
            seg[2] = item.Cell;
        }
    }
}
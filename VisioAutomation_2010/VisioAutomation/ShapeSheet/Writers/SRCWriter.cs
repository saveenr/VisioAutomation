using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SRCWriter : WriterBase<SRC>
    {
        public SRCWriter() : base()
        {
        }

        public SRCWriter(int capacity) : base(capacity)
        {
        }

        protected override short[] build_stream()
        {
            var streamb = new List<SRC>(this._updates.Count);
            streamb.AddRange(this._updates.Select(i => i.StreamItem));
            return SRC.ToStream(streamb);
        }
    }
}
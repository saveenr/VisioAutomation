using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class SIDSRCWriter : WriterBase<SIDSRC>
    {
        public SIDSRCWriter() : base()
        {
        }

        public SIDSRCWriter(int capacity) : base(capacity)
        {
        }

        protected override short[] build_stream()
        {
            var streamb = new List<SIDSRC>(this._updates.Count);
            streamb.AddRange(this._updates.Select(i => i.StreamItem));
            return SIDSRC.ToStream(streamb);
        }
    }
}
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public struct SidSrc
    {
        public short ShapeID { get; }
        public Src Src { get; }

        public SidSrc(
            short shape_id,
            Src src)
        {
            this.ShapeID = shape_id;
            this.Src = src;
        }  
        
        public override string ToString()
        {
            return string.Format("{0}({1},{2},{3},{4})", nameof(SidSrc),this.ShapeID, this.Src.Section, this.Src.Row, this.Src.Cell);
        }

        public static short [] ToStream(IList<SidSrc> sidsrcs)
        {
            const int sidsrc_length = 4;
            var sidsrcstream = new short[sidsrc_length*sidsrcs.Count];
            for (int i = 0; i < sidsrcs.Count; i++)
            {
                var sidsrc = sidsrcs[i];
                int pos = i*sidsrc_length;
                sidsrcstream[pos + 0] = sidsrc.ShapeID;
                sidsrcstream[pos + 1] = sidsrc.Src.Section;
                sidsrcstream[pos + 2] = sidsrc.Src.Row;
                sidsrcstream[pos + 3] = sidsrc.Src.Cell;
            }
            return sidsrcstream;
        }
    }
}
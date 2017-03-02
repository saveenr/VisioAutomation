using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct SidSrc
    {
        public short ShapeID { get; }
        public short Section { get; }
        public short Row { get; }
        public short Cell { get; }

        public SidSrc(
            short shape_id,
            IVisio.VisSectionIndices section,
            IVisio.VisRowIndices row,
            IVisio.VisCellIndices cell) : this(shape_id,(short)section,(short)row,(short)cell)
        {
        }

        public SidSrc(
            short shape_id,
            short section,
            short row,
            short cell) : this()
        {
            this.ShapeID = shape_id;
            this.Section = section;
            this.Row = row;
            this.Cell = cell;
        }

        public SidSrc(
            short shape_id,
            Src src) : this(shape_id,src.Section,src.Row,src.Cell)
        {
        }  
        
        public override string ToString()
        {
            return string.Format("{0}({1},{2},{3},{4})", nameof(SidSrc),this.ShapeID, this.Section, this.Row, this.Cell);
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
                sidsrcstream[pos + 1] = sidsrc.Section;
                sidsrcstream[pos + 2] = sidsrc.Row;
                sidsrcstream[pos + 3] = sidsrc.Cell;
            }
            return sidsrcstream;
        }

        public Src SRC
        {
            get { return new Src(this.Section, this.Row, this.Cell); }
        }
    }
}
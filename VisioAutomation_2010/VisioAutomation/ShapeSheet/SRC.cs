using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct SRC
    {
        public short Section { get; }
        public short Row { get; }
        public short Cell { get; }

        public SRC(
            IVisio.VisSectionIndices section,
            IVisio.VisRowIndices row,
            IVisio.VisCellIndices cell)
            : this((short)section, (short)row, (short)cell)
        {
        }

        public SRC(
            short section,
            short row,
            short cell)
            : this()
        {
            this.Section = section;
            this.Row = row;
            this.Cell = cell;
        }

        public override string ToString()
        {
            return string.Format("{0}({1},{2},{3})", nameof(SRC), this.Section, this.Row, this.Cell);
        }

        public SRC WithRow(short row)
        {
            // It's common to need to get a SRC that has a different row index.
            // This method make that very easy
            return new SRC(this.Section, row, this.Cell);
        }

        public static short[] ToStream(IList<SRC> srcs)
        {
            const int src_length = 3;
            var s = new short[src_length * srcs.Count];
            for (int i = 0; i < srcs.Count; i++)
            {
                var sidsrc = srcs[i];
                int pos = i * src_length;
                s[pos + 0] = sidsrc.Section;
                s[pos + 1] = sidsrc.Row;
                s[pos + 2] = sidsrc.Cell;
            }
            return s;
        }

    }
}

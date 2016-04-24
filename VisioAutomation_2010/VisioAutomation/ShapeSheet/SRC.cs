using System;
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
            return String.Format("({0},{1},{2})", this.Section, this.Row, this.Cell);
        }

        public SRC ForRow(short row) => new SRC(this.Section, row, this.Cell);

        public SRC ForSectionAndRow(short section, short row) => new SRC(section, row, this.Cell);

        public bool AreEqual(SRC other)
        {
            return ((this.Section == other.Section) && (this.Row == other.Row) && (this.Cell == other.Cell));
        }

        internal delegate SRC SRCFromCellIndex(IVisio.VisCellIndices c);

        internal static SRCFromCellIndex GetSRCFactory(IVisio.VisSectionIndices sec, IVisio.VisRowIndices row)
        {
            SRCFromCellIndex new_func = cell => new SRC(sec, row, cell);
            return new_func;
        }

        public static short[] ToStream(IList<SRC> srcs)
        {
            var s = new short[3 * srcs.Count];
            for (int i = 0; i < srcs.Count; i++)
            {
                var sidsrc = srcs[i];
                int pos = i * 3;
                s[pos + 0] = sidsrc.Section;
                s[pos + 1] = sidsrc.Row;
                s[pos + 2] = sidsrc.Cell;
            }
            return s;
        }

    }
}

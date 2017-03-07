using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public struct Src
    {
        public short Section { get; }
        public short Row { get; }
        public short Cell { get; }

        public Src(
            IVisio.VisSectionIndices section,
            IVisio.VisRowIndices row,
            IVisio.VisCellIndices cell)
            : this((short)section, (short)row, (short)cell)
        {
        }

        public Src(
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
            return string.Format("{0}({1},{2},{3})", nameof(Src), this.Section, this.Row, this.Cell);
        }

        public Src CloneWithNewRow(short row)
        {
            // Src that has a different row index. Very common scenario
            return new Src(this.Section, row, this.Cell);
        }
    }
}

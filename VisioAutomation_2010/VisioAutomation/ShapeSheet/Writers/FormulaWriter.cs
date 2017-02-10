using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class FormulaWriter : WriterBase<FormulaLiteral>
    {
        public FormulaWriter() :base()
        {
        }

        public void SetFormula(SRC src, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.Add(src,formula);
            }
        }

        public void SetFormula(short id, SRC src, FormulaLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormula(SIDSRC sidsrc, FormulaLiteral formula)
        {
            this.__SetFormulaIgnoreNull(sidsrc, formula);
        }

        private void __SetFormulaIgnoreNull(SIDSRC sidsrc, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.Add(sidsrc,formula);
            }
        }

        protected override void CommitRecordsByType(ShapeSheetSurface surface, CoordType coord_type)
        {
            var records = this.GetRecords(coord_type);
            var count = records.Count();

            if (count == 0)
            {
                return;
            }

            int chunksize = coord_type == CoordType.SIDSRC ? 4 : 3;

            var stream = new short[count * chunksize];
            var formulas = new object[count];

            int streampos = 0;
            int formulapos = 0;

            foreach (var rec in records)
            {
                // fill stream
                if (coord_type == CoordType.SRC)
                {
                    var src = rec.SRC;
                    stream[streampos++] = src.Section;
                    stream[streampos++] = src.Row;
                    stream[streampos++] = src.Cell;
                }
                else
                {
                    var sidsrc = rec.SIDSRC;
                    stream[streampos++] = sidsrc.ShapeID;
                    stream[streampos++] = sidsrc.Section;
                    stream[streampos++] = sidsrc.Row;
                    stream[streampos++] = sidsrc.Cell;
                }

                // fill formulas
                formulas[formulapos++] = rec.Value.Value;
            }

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }
    }
}
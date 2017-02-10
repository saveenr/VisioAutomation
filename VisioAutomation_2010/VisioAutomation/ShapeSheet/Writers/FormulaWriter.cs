using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class FormulaWriter : XWriterBase<FormulaLiteral>
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

        protected override void CommitSIDSRC(ShapeSheetSurface surface)
        {
            var sidsrc_records = this.GetSIDSRCRecords();
            var count = sidsrc_records.Count();

            if (count == 0)
            {
                return;
            }

            var stream = new short[count * 4];
            var formulas = new object[count];

            int streampos = 0;
            int formulapos= 0;

            foreach (var rec in sidsrc_records)
            {
                // fill stream
                var sidsrc = rec.Sidsrc;
                stream[streampos++] = sidsrc.ShapeID;
                stream[streampos++] = sidsrc.Section;
                stream[streampos++] = sidsrc.Row;
                stream[streampos++] = sidsrc.Cell;

                // fill formulas
                formulas[formulapos++] = rec.Value.Value;
            }

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        protected override void CommitSRC(ShapeSheetSurface surface)
        {
            var srcrecords = this.GetSRCRecords();
            var count = srcrecords.Count();

            if (count == 0)
            {
                return;
            }

            var stream = new short[count * 3];
            var formulas = new object[count];

            int streampos = 0;
            int formulapos = 0;

            foreach (var rec in srcrecords)
            {
                // fill stream
                var src = rec.SRC;
                stream[streampos++] = src.Section;
                stream[streampos++] = src.Row;
                stream[streampos++] = src.Cell;

                // fill formulas
                formulas[formulapos++] = rec.Value.Value;
            }

            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }
    }
}
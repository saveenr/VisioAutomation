using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class FormulaWriterSIDSRC : WriterBase<VisioAutomation.ShapeSheet.SIDSRC,FormulaLiteral>
    {
        public FormulaWriterSIDSRC() : base()
        {
        }

        public FormulaWriterSIDSRC(int capacity) : base(capacity)
        {
        }

        public void SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this.StreamItems.Add(streamitem);
            this.ValueItems.Add(formula);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this.StreamItems.Add(streamitem);
            this.ValueItems.Add(formula);
        }

        public void SetFormulaIgnoreNull(short id, SRC src, FormulaLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this._SetFormulaIgnoreNull(sidsrc, formula);
        }

        public void SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(streamitem, formula);
        }

        protected void _SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.SetFormula(streamitem,formula);
            }
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SIDSRC.ToStream(this.StreamItems);
            var formulas = WriterHelper.build_formulas_array(this.ValueItems);
            var flags = this.GetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }
    }
}
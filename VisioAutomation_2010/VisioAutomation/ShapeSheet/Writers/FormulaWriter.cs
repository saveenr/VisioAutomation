using System.Collections.Generic;

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

        protected override void CommitSIDSRC(ShapeSheetSurface surface)
        {
            var stream = this.GetSIDSRCStream();
            var formulas = build_formulas_array(this.SIDSRC_Values);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        protected override void CommitSRC(ShapeSheetSurface surface)
        {
            var stream = this.GetSRCStream();
            var formulas = build_formulas_array(this.SRC_Values);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        private static object[] build_formulas_array(IList<FormulaLiteral> formulas)
        {
            var result = new object[formulas.Count];
            int i = 0;
            foreach (var rec in formulas)
            {
                result[i] = rec.Value;
                i++;
            }
            return result;
        }
    }
}
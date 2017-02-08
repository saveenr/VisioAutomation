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

        public override void Commit(ShapeSheetSurface surface)
        {
            this.CommitSRC(surface);
            this.CommitSIDSRC(surface);            
        }

        protected void CommitSIDSRC(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SIDSRCCount < 1)
            {
                return;
            }

            var stream = this.GetSIDSRCStream();
            var formulas = WriterHelper.build_formulas_array(this.SIDSRC_Values);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        protected void CommitSRC(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SRCCount < 1)
            {
                return;
            }

            var stream = this.GetSRCStream();
            var formulas = WriterHelper.build_formulas_array(this.SRC_Values);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }
    }
}
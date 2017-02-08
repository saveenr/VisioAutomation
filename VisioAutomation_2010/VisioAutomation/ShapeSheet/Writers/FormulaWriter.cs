namespace VisioAutomation.ShapeSheet.Writers
{
    public class FormulaWriter : WriterBaseEx<FormulaLiteral>
    {
        public FormulaWriter() :base()
        {
        }

        public void SetFormula(SRC src, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.SRC_StreamItems.Add(src);
                this.SRC_ValueItems.Add(formula);
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

        protected void __SetFormulaIgnoreNull(SIDSRC sidsrc, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.SIDSRC_StreamItems.Add(sidsrc);
                this.SIDSRC_ValueItems.Add(formula);
            }
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            this.SRC_commit_to_surface(surface);
            this.SIDSRC_commit_to_surface(surface);            
        }

        protected void SIDSRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SIDSRC_ValueItems.Count < 1)
            {
                return;
            }

            var stream = SIDSRC.ToStream(this.SIDSRC_StreamItems);
            var formulas = WriterHelper.build_formulas_array(this.SIDSRC_ValueItems);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

        protected void SRC_commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.SRC_ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.SRC_StreamItems);
            var formulas = WriterHelper.build_formulas_array(this.SRC_ValueItems);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }
    }
}
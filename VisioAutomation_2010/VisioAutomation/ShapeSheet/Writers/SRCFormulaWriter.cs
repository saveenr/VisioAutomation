namespace VisioAutomation.ShapeSheet.Writers
{
    public class SRCFormulaWriter : WriterBase<VisioAutomation.ShapeSheet.SRC, FormulaLiteral>
    {
        public SRCFormulaWriter() :base()
        {
        }

        public SRCFormulaWriter(int capacity) : base( capacity )
        {
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this.StreamItems.Add(streamitem);
            this.ValueItems.Add(formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.SetFormula(streamitem, formula);
            }
        }

        public override void Commit(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.StreamItems);
            var formulas = WriterBase<VisioAutomation.ShapeSheet.SRC, FormulaLiteral>.build_formulas(this.ValueItems);
            var flags = this.FormulaFlags;
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

    }
}
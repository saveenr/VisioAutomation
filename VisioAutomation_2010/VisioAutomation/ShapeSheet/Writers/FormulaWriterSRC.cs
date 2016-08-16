namespace VisioAutomation.ShapeSheet.Writers
{
    public class FormulaWriterSRC : WriterBase<VisioAutomation.ShapeSheet.SRC, FormulaLiteral>
    {
        public FormulaWriterSRC() :base()
        {
        }

        public FormulaWriterSRC(int capacity) : base( capacity )
        {
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this.StreamItems.Add(streamitem);
                this.ValueItems.Add(formula);
            }
        }

        protected override void _commit_to_surface(ShapeSheetSurface surface)
        {
            // Do nothing if there aren't any updates
            if (this.ValueItems.Count < 1)
            {
                return;
            }

            var stream = SRC.ToStream(this.StreamItems);
            var formulas = WriterHelper.build_formulas_array(this.ValueItems);
            var flags = this.ComputeGetFormulaFlags();
            int c = surface.SetFormulas(stream, formulas, (short)flags);
        }

    }
}
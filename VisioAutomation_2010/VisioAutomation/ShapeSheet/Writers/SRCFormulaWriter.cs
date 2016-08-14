namespace VisioAutomation.ShapeSheet.Writers
{
    public class SRCFormulaWriter : SRCWriter
    {
        public SRCFormulaWriter() :base()
        {
        }

        public SRCFormulaWriter(int capacity) : base( capacity )
        {
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(streamitem, formula);
        }

        protected void _SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new WriterRecord<SRC>(streamitem, formula.Value);
            this._add_update(rec);
        }
        protected void _SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this._SetFormula(streamitem, formula);
            }
        }
    }
}
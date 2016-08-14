namespace VisioAutomation.ShapeSheet.Writers
{
    public class SIDSRCFormulaWriter : SIDSRCWriter
    {
        public SIDSRCFormulaWriter() : base()
        {
        }

        public SIDSRCFormulaWriter(int capacity) : base(capacity)
        {
        }

        protected void _SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this.CheckFormulaIsNotNull(formula.Value);
            var rec = new WriterRecord<SIDSRC>(streamitem, formula.Value);
            this._add_update(rec);
        }

        public void SetFormula(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(streamitem, formula);
        }

        public void SetFormula(short shapeid, SRC src, FormulaLiteral formula)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetFormula(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(streamitem, formula);
        }

        public void SetFormulaIgnoreNull(short id, SRC src, FormulaLiteral formula)
        {
            var sidsrc = new SIDSRC(id, src);
            this._SetFormulaIgnoreNull(sidsrc, formula);
        }

        protected void _SetFormulaIgnoreNull(SIDSRC streamitem, FormulaLiteral formula)
        {
            if (formula.HasValue)
            {
                this._SetFormula(streamitem, formula);
            }
        }

    }
}
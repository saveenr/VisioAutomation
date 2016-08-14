using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateSIDSRC : UpdateBase
    {
        public UpdateSIDSRC() : base()
        {
        }

        public UpdateSIDSRC(int capacity) : base(capacity)
        {
        }

        public void SetResult(short shapeid, SRC src, double value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(short shapeid, SRC src, string value, IVisio.VisUnitCodes unitcode)
        {
            var streamitem = new SIDSRC(shapeid, src);
            this._SetResult(streamitem, value, unitcode);
        }

        public void SetResult(SIDSRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(streamitem, value, unitcode);
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

    }
}
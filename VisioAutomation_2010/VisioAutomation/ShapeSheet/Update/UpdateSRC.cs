using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public class UpdateSRC : UpdateBase
    {
        public UpdateSRC() :base()
        {
        }

        public UpdateSRC(int capacity) : base( capacity )
        {
        }

        public void SetFormula(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormula(StreamType.SRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetFormulaIgnoreNull(SRC streamitem, FormulaLiteral formula)
        {
            this._SetFormulaIgnoreNull(StreamType.SRC, new SIDSRC(-1, streamitem), formula);
        }

        public void SetResult(SRC streamitem, string value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SRC, new SIDSRC(-1, streamitem), value, unitcode);
        }

        public void SetResult(SRC streamitem, double value, IVisio.VisUnitCodes unitcode)
        {
            this._SetResult(StreamType.SRC, new SIDSRC(-1, streamitem), value, unitcode);
        }

    }
}
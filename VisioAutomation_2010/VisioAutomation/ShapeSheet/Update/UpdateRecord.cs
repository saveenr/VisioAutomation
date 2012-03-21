using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public struct UpdateRecord 
    {
        public readonly SIDSRC SIDSRC;
        public readonly string Formula;
        public readonly double Result;
        public readonly IVisio.VisUnitCodes UnitCode;
        public readonly UpdateType UpdateType;

        internal UpdateRecord(SIDSRC sidsrc, string formula)
        {
            this.SIDSRC = sidsrc;
            this.Formula = formula;
            this.Result = 0.0;
            this.UnitCode = IVisio.VisUnitCodes.visNoCast;
            this.UpdateType  = UpdateType.Formula;
        }

        internal UpdateRecord(SIDSRC sidsrc, double result, IVisio.VisUnitCodes unit_code)
        {
            this.SIDSRC = sidsrc;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.Result = result;
            this.UpdateType = UpdateType.Result;
        }
    }
}
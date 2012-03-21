using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public struct UpdateRecord 
    {
        public readonly SIDSRC StreamItem;
        public readonly string Formula;
        public readonly double Result;
        public readonly IVisio.VisUnitCodes UnitCode;
        public readonly UpdateType UpdateType;

        public UpdateRecord(SIDSRC streamitem, string formula)
        {
            this.StreamItem = streamitem;
            this.Formula = formula;
            this.Result = 0.0;
            this.UnitCode = IVisio.VisUnitCodes.visNoCast;
            this.UpdateType  = UpdateType.Formula;
        }

        public UpdateRecord(SIDSRC streamitem, double result, IVisio.VisUnitCodes unit_code)
        {
            this.StreamItem = streamitem;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.Result = result;
            this.UpdateType = UpdateType.Result;
        }
    }
}
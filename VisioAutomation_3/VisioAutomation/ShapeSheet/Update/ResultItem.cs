using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Update
{
    public struct ResultItem<TStream> where TStream : struct
    {
        public readonly TStream StreamItem;
        public readonly IVisio.VisUnitCodes UnitCode;
        public readonly double Result;

        public ResultItem(TStream streamitem, double result, IVisio.VisUnitCodes unit_code)
        {
            this.StreamItem = streamitem;
            this.UnitCode = unit_code;
            this.Result = result;
        }
    }
}
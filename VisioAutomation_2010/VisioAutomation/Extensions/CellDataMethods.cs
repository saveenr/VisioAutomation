using VisioAutomation.Utilities;

namespace VisioAutomation.Extensions
{
    public static class CellDataMethods
    {
        public static ShapeSheet.CellData ToInt(this ShapeSheet.CellData celldata)
        {
            return new ShapeSheet.CellData(celldata.Formula,celldata.Result);
        }

        public static ShapeSheet.CellData ToBool(this ShapeSheet.CellData celldata)
        {
            return new ShapeSheet.CellData(celldata.Formula,celldata.Result);
        }
    }
}
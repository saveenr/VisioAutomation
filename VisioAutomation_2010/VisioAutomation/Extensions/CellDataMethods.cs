namespace VisioAutomation.Extensions
{
    public static class CellDataMethods
    {
        public static ShapeSheet.CellData<int> ToInt(this ShapeSheet.CellData<double> celldata)
        {
            return new ShapeSheet.CellData<int>(celldata.Formula,(int)celldata.Result);
        }

        public static ShapeSheet.CellData<bool> ToBool(this ShapeSheet.CellData<double> celldata)
        {
            return new ShapeSheet.CellData<bool>(celldata.Formula, Convert.DoubleToBool(celldata.Result));
        }
    }
}
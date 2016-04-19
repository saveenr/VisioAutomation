namespace VisioAutomation.Extensions
{
    public static class CellDataMethods
    {
        public static ShapeSheet.CellData<int> ToInt(this ShapeSheet.CellData<double> cd)
        {
            return new ShapeSheet.CellData<int>(cd.Formula,(int)cd.Result);
        }

        public static ShapeSheet.CellData<bool> ToBool(this ShapeSheet.CellData<double> cd)
        {
            return new ShapeSheet.CellData<bool>(cd.Formula, Convert.DoubleToBool(cd.Result));
        }
    }
}
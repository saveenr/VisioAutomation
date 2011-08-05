using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class CellDataExtensions
    {
        public static VA.ShapeSheet.CellData<int> ToInt(this VA.ShapeSheet.CellData<double> cd)
        {
            return cd.Cast(v => (int) v);
        }

        public static VA.ShapeSheet.CellData<bool> ToBool(this VA.ShapeSheet.CellData<double> cd)
        {
            return cd.Cast(v => VA.Convert.DoubleToBool(v));
        }
    }
}
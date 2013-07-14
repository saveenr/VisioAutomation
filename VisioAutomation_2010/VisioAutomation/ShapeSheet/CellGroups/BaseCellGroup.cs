using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        public delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        public delegate T RowToCells<T>(CellData<double>[] data);
        public delegate T ResultToCells<T>(VA.ShapeSheet.Query.CellQuery.QueryResult<CellData<double>>  data);
    }
}
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using TABLE = VisioAutomation.ShapeSheet.Data.Table<VisioAutomation.ShapeSheet.CellData<double>>;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        // Delegates
        public delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected delegate TObj RowToCells<TQuery, TObj>(TQuery query, TABLE table, int row) where TQuery : VA.ShapeSheet.Query.QueryBase;
    }
}
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using TABLEROW = VisioAutomation.ShapeSheet.Data.TableRow<VisioAutomation.ShapeSheet.CellData<double>>;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class BaseCellGroup
    {
        // Delegates
        protected delegate void ApplyFormula(VA.ShapeSheet.SRC src, VA.ShapeSheet.FormulaLiteral formula);
        protected delegate TObj RowToCells<TQuery, TObj>(TQuery query, TABLEROW tablerow) where TQuery : VA.ShapeSheet.Query.QueryBase;
    }
}
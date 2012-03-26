using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using TABLEROW = VisioAutomation.ShapeSheet.Data.TableRow<VisioAutomation.ShapeSheet.CellData<double>>;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        // Delegates
        protected delegate TObj RowToCells<TQuery, TObj>(TQuery query, TABLEROW tablerow) where TQuery : VA.ShapeSheet.Query.CellQuery;
        
        protected abstract void ApplyFormulas(ApplyFormula func);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this.ApplyFormulas((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var cells = table.Select(r => row_to_cells_func(query, r));
            var cells_list = new List<TObj>(table.RowCount);
            cells_list.AddRange(cells);
            return cells_list;
        }

        protected static TObj CellsFromRow<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            var tablerow = table[0];
            return row_to_obj_func(query, tablerow);
        }
    }
}
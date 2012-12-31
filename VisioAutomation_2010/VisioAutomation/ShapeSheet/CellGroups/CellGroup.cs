using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {       
        public abstract void ApplyFormulas(ApplyFormula func);

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var cells_list = new List<TObj>(table.RowCount);
            for (int i = 0; i < table.RowCount; i++)
            {
                var new_object = row_to_cells_func(query, table, i);
                cells_list.Add(new_object);
            }
            return cells_list;
        }

        protected static TObj CellsFromRow<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            return row_to_obj_func(query, table, 0);
        }
    }
}
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        // Delegates
        protected delegate TObj RowToObject<TQuery, TObj>(TQuery query, VA.ShapeSheet.Data.QueryDataRow<double> qdr) where TQuery : VA.ShapeSheet.Query.CellQuery;
        
        protected abstract void _Apply(ApplyFormula func);

        public void Apply(VA.ShapeSheet.Update.SIDSRCUpdate update, short shapeid)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(shapeid, src, f));
        }

        public void Apply(VA.ShapeSheet.Update.SRCUpdate update)
        {
            this._Apply((src, f) => update.SetFormulaIgnoreNull(src, f));
        }

        protected static IList<TObj> _GetObjectsFromRows<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToObject<TQuery, TObj> row_to_cells_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(page, shapeids);
            var rows = qds.EnumRows();
            var objs = rows.Select(r => row_to_cells_func(query, r));
            var obj_list = new List<TObj>(qds.RowCount);
            obj_list.AddRange(objs);
            return obj_list;
        }

        protected static TObj _GetObjectFromSingleRow<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToObject<TQuery, TObj> row_to_obj_func) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var qds = query.GetFormulasAndResults<double>(shape);
            var qdr = qds.GetRow(0);
            return row_to_obj_func(query, qdr);
        }
    }
}
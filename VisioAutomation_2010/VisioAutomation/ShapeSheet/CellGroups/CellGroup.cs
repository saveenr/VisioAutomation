using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {       
        public abstract void ApplyFormulas(ApplyFormula func);

        protected static IList<TObj> CellsFromRows<TQuery, TObj>(IVisio.Page page, IList<int> shapeids, TQuery query, RowToCells<TQuery, TObj> row_to_cells) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(page, shapeids);
            var objects = new List<TObj>(table.RowCount);
            for (int i = 0; i < table.RowCount; i++)
            {
                var new_object = row_to_cells(query, table, i);
                objects.Add(new_object);
            }
            return objects;
        }

        protected static TObj CellsFromRow<TQuery, TObj>(IVisio.Shape shape, TQuery query, RowToCells<TQuery, TObj> row_to_obj) where TQuery : VA.ShapeSheet.Query.CellQuery
        {
            var table = query.GetFormulasAndResults<double>(shape);
            return row_to_obj(query, table, 0);
        }
    }

    public abstract class CellGroupEx : CellGroup
    {
        protected static IList<T> _GetCells<T>(IVisio.Page page, IList<int> shapeids, QueryEx query, System.Func<ExQueryResult<CellData<double>>,T> f )
        {
            var data = query.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<T>();
            for (int i = 0; i < shapeids.Count; i++)
            {
                var cells = f(data[i]);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T>(IVisio.Shape shape, QueryEx query, System.Func<ExQueryResult<CellData<double>>, T> f)
        {
            var data_for_shape = query.GetFormulasAndResults<double>(shape);
            var cells = f(data_for_shape);
            return cells;
        }
    }
}
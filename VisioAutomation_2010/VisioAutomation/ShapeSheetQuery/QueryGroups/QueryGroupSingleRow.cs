using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheetQuery.Outputs;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class QueryGroupSingleRow : QueryGroupBase
    {
        private static void verify_cell_only_query(Query query)
        {
            if (query.Cells.Count < 1)
            {
                throw new AutomationException("Query must contain at least 1 Column");
            }

            if (query.SubQueries.Count != 0)
            {
                throw new AutomationException("Query should not contain contain any sections");
            }
        }

        protected static IList<T> _GetCells<T, TResult>(
            IVisio.Page page, IList<int> shapeids,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cells_to_object)
        {
            verify_cell_only_query(query);

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>( surface, shapeids);
            var list = new List<T>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var cells = cells_to_object(data_for_shape.Cells);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T, TResult>(
            IVisio.Shape shape,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cells_to_object)
        {
            verify_cell_only_query(query);

            var ss1 = new ShapeSheetSurface(shape);
            Output<CellData<TResult>> data_for_shape = query.GetFormulasAndResults<TResult>(ss1);
            var cells = cells_to_object(data_for_shape.Cells);
            return cells;
        }

        public void SetFormulas(VisioAutomation.ShapeSheet.Update.UpdateSRCFormulas update)
        {
            foreach (var pair in this.Pairs)
            {
                update.SetFormulaIgnoreNull(pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, VisioAutomation.ShapeSheet.Update.UpdateSIDSRCFormula update)
        {
            foreach (var pair in this.Pairs)
            {
                update.SetFormulaIgnoreNull(shapeid, pair.SRC, pair.Formula);
            }
        }
    }
}
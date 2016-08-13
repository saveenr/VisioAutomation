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

        protected static IList<T> _GetCells<T, RT>(
            IVisio.Page page, IList<int> shapeids,
            Query query,
            CellsToObject<T, RT> cell_data_to_object)
        {
            verify_cell_only_query(query);

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetCellData<RT>( surface, shapeids);
            var list = new List<T>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var cells = cell_data_to_object(data_for_shape.Cells);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T, RT>(
            IVisio.Shape shape,
            Query query,
            CellsToObject<T, RT> cell_data_to_object)
        {
            verify_cell_only_query(query);

            var ss1 = new ShapeSheetSurface(shape);
            Output<CellData<RT>> data_for_shape = query.GetCellData<RT>(ss1);
            var cells = cell_data_to_object(data_for_shape.Cells);
            return cells;
        }
    }
}
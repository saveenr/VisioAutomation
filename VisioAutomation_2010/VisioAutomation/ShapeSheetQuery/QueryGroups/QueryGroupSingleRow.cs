using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class QueryGroupSingleRow : QueryGroupBase
    {
        private static void check_query(CellQuery query)
        {
            if (query.CellColumns.Count < 1)
            {
                throw new AutomationException("Query must contain at least 1 Column");
            }

            if (query.SectionColumns.Count != 0)
            {
                throw new AutomationException("Query should not contain contain any sections");
            }
        }

        protected static IList<T> _GetCells<T, RT>(
            IVisio.Page page, IList<int> shapeids,
            CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            var surface = new ShapeSheet.ShapeSheetSurface(page);
            var data_for_shapes = query.GetCellData<RT>( surface, shapeids);
            var list = new List<T>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var cells = row_to_object(data_for_shape.Cells);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T, RT>(
            IVisio.Shape shape,
            CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            QueryResult<CellData<RT>> data_for_shape = query.GetCellData<RT>(shape);
            var cells = row_to_object(data_for_shape.Cells);
            return cells;
        }
    }
}
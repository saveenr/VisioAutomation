using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        private static void check_query(VA.ShapeSheet.Query.CellQuery query)
        {
            if (query.CellColumns.Count < 1)
            {
                throw new VA.AutomationException("Query must contain at least 1 Column");
            }

            if (query.SectionColumns.Count != 0)
            {
                throw new VA.AutomationException("Query should not contain contain any sections");
            }
        }

        protected static IList<T> _GetCells<T, RT>(
            IVisio.Page page, IList<int> shapeids,
            VA.ShapeSheet.Query.CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            var surface = new ShapeSheetSurface(page);
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
            VA.ShapeSheet.Query.CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            var data_for_shape = query.GetCellData<RT>(shape);
            var cells = row_to_object(data_for_shape.Cells);
            return cells;
        }
    }
}
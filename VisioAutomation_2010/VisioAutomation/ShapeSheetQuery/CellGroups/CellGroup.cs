using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        private static void check_query(VisioAutomation.ShapeSheetQuery.CellQuery query)
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
            VisioAutomation.ShapeSheetQuery.CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            CellGroup.check_query(query);

            var surface = new VisioAutomation.ShapeSheet.ShapeSheetSurface(page);
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
            VisioAutomation.ShapeSheetQuery.CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            CellGroup.check_query(query);

            var data_for_shape = query.GetCellData<RT>(shape);
            var cells = row_to_object(data_for_shape.Cells);
            return cells;
        }
    }
}
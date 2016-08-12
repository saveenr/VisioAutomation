using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheetQuery.Results;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class CellQueryGroupMultiRow : QueryGroupBase
    {
        private static void check_query(Query query)
        {
            if (query.Cells.Count != 0)
            {
                throw new AutomationException("Query should not contain any Columns");
            }

            if (query.SubQueries.Count != 1)
            {
                throw new AutomationException("Query should not contain contain exaxtly 1 section");
            }
        }


        public static IList<List<T>> _GetCells<T, RT>(
            IVisio.Page page,
            IList<int> shapeids,
            Query query,
            RowToObject<T, RT> row_to_object)
        {
            CellQueryGroupMultiRow.check_query(query);

            var list = new List<List<T>>(shapeids.Count);
            var surface = new QuerySurface(page);
            var data_for_shapes = query.GetCellData<RT>(surface, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var sec = data_for_shape.Sections[0];
                var sec_objects = CellQueryGroupMultiRow.SectionRowsToObjects(sec, row_to_object);
                list.Add(sec_objects);
            }

            return list;
        }

        public static IList<T> _GetCells<T, RT>(
            IVisio.Shape shape,
            Query query,
            RowToObject<T, RT> row_to_object)
        {
            CellQueryGroupMultiRow.check_query(query);

            var ss1 = new QuerySurface(shape);
            var data_for_shape = query.GetCellData<RT>(ss1);
            var sec = data_for_shape.Sections[0];
            var sec_objects = CellQueryGroupMultiRow.SectionRowsToObjects(sec, row_to_object);
            
            return sec_objects;
        }

        private static List<T> SectionRowsToObjects<T, RT>(SectionSubQueryResult<ShapeSheet.CellData<RT>> sec, RowToObject<T, RT> row_to_object)
        {
            int num_rows = sec.Rows.Count;
            var sec_objects = new List<T>(num_rows);
            foreach (var row in sec.Rows)
            {
                var obj = row_to_object(row.Cells);
                sec_objects.Add(obj);
            }
            return sec_objects;
        }
    }
}
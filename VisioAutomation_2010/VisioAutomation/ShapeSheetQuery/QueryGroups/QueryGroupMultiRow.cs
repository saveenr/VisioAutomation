using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheetQuery.Results;

namespace VisioAutomation.ShapeSheetQuery.QueryGroups
{
    public abstract class QueryGroupMultiRow : QueryGroupBase
    {
        private static void verify_single_section_query(Query query)
        {
            if (query.Cells.Count != 0)
            {
                throw new AutomationException("Query should not contain any Columns");
            }

            if (query.SubQueries.Count != 1)
            {
                throw new AutomationException("Query should not contain contain exactly 1 section");
            }
        }


        public static IList<List<T>> _GetCells<T, RT>(
            IVisio.Page page,
            IList<int> shapeids,
            Query query,
            RowToObject<T, RT> row_to_object)
        {
            QueryGroupMultiRow.verify_single_section_query(query);

            var list = new List<List<T>>(shapeids.Count);
            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetCellData<RT>(surface, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var sec = data_for_shape.Sections[0];
                var sec_objects = QueryGroupMultiRow.SectionRowsToObjects(sec, row_to_object);
                list.Add(sec_objects);
            }

            return list;
        }

        public static IList<T> _GetCells<T, RT>(
            IVisio.Shape shape,
            Query query,
            RowToObject<T, RT> row_to_object)
        {
            QueryGroupMultiRow.verify_single_section_query(query);

            var ss1 = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetCellData<RT>(ss1);
            var sec = data_for_shape.Sections[0];
            var sec_objects = QueryGroupMultiRow.SectionRowsToObjects(sec, row_to_object);
            
            return sec_objects;
        }

        private static List<T> SectionRowsToObjects<T, RT>(SubQueryResult<ShapeSheet.CellData<RT>> sec, RowToObject<T, RT> row_to_object)
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
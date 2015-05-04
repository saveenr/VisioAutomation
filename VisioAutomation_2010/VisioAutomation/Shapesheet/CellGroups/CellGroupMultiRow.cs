using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        private static void check_query(CellQuery query)
        {
            if (query.CellColumns.Count != 0)
            {
                throw new AutomationException("Query should not contain any Columns");
            }

            if (query.SectionColumns.Count != 1)
            {
                throw new AutomationException("Query should not contain contain exaxtly 1 section");
            }
        }


        public static IList<List<T>> _GetCells<T, RT>(
            IVisio.Page page,
            IList<int> shapeids,
            CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            var list = new List<List<T>>(shapeids.Count);
            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetCellData<RT>(surface, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var sec = data_for_shape.Sections[0];
                var sec_objects = SectionToObjectList(sec, row_to_object);
                list.Add(sec_objects);
            }

            return list;
        }

        public static IList<T> _GetCells<T, RT>(
            IVisio.Shape shape,
            CellQuery query,
            RowToObject<T, RT> row_to_object)
        {
            check_query(query);

            var data_for_shape = query.GetCellData<RT>(shape);
            var sec = data_for_shape.Sections[0];
            var sec_objects = SectionToObjectList(sec, row_to_object);
            
            return sec_objects;
        }

        private static List<T> SectionToObjectList<T, RT>(SectionResult<CellData<RT>> sec, RowToObject<T, RT> row_to_object)
        {
            int num_rows = sec.Count;
            var sec_objects = new List<T>(num_rows);
            foreach (var row in sec)
            {
                var obj = row_to_object(row);
                sec_objects.Add(obj);
            }
            return sec_objects;
        }
    }
}
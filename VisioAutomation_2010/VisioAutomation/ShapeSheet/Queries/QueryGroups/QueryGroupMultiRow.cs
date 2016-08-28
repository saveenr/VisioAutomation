using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Queries.Outputs;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.Queries.QueryGroups
{
    public abstract class QueryGroupMultiRow : QueryGroupBase
    {
        private static void verify_single_section_query(Query query)
        {
            if (query.Cells.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain any Columns");
            }

            if (query.SubQueries.Count != 1)
            {
                throw new InternalAssertionException("Query should not contain contain exactly 1 section");
            }
        }


        public static IList<List<T>> _GetCells<T, TResult>(
            IVisio.Page page,
            IList<int> shapeids,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cell_data_to_object)
        {
            QueryGroupMultiRow.verify_single_section_query(query);

            var list = new List<List<T>>(shapeids.Count);
            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>(surface, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var sec = data_for_shape.Sections[0];
                var sec_objects = QueryGroupMultiRow.SectionRowsToObjects(sec, cell_data_to_object);
                list.Add(sec_objects);
            }

            return list;
        }

        public static IList<T> _GetCells<T, TResult>(
            IVisio.Shape shape,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cell_data_to_object)
        {
            QueryGroupMultiRow.verify_single_section_query(query);

            var ss1 = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetFormulasAndResults<TResult>(ss1);
            var sec = data_for_shape.Sections[0];
            var sec_objects = QueryGroupMultiRow.SectionRowsToObjects(sec, cell_data_to_object);
            
            return sec_objects;
        }

        private static List<T> SectionRowsToObjects<T, TResult>(SubQueryOutput<ShapeSheet.CellData<TResult>> sec, System.Func<ShapeSheet.CellData<TResult>[],T> cells_to_object)
        {
            int num_rows = sec.Rows.Count;
            var sec_objects = new List<T>(num_rows);
            foreach (var row in sec.Rows)
            {
                var obj = cells_to_object(row.Cells);
                sec_objects.Add(obj);
            }
            return sec_objects;
        }

        public void SetFormulas(short shapeid, FormulaWriterSIDSRC writer,
            short row)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(shapeid, pair.SRC.CopyWithNewRow(row), pair.Formula);
            }
        }

        public void SetFormulas(FormulaWriterSRC writer, short row)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(pair.SRC.CopyWithNewRow(row), pair.Formula);
            }
        }

    }
}
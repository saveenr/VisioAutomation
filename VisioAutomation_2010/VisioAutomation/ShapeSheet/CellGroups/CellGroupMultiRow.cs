using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Queries;
using VisioAutomation.ShapeSheet.Queries.Outputs;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : CellGroupBase
    {
        private static void verify_multirow_query(Query query)
        {
            if (query.Cells.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain any cells");
            }

            if (query.SubQueries.Count != 1)
            {
                throw new InternalAssertionException("Query should contain contain exactly 1 subquery");
            }
        }


        protected static IList<List<T>> _GetCells<T, TResult>(
            IVisio.Page page,
            IList<int> shapeids,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cell_data_to_object)
        {
            CellGroupMultiRow.verify_multirow_query(query);

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<List<T>>(shapeids.Count);
            var objects = data_for_shapes.Select(d => CellGroupMultiRow.SectionRowsToObjects(d.Sections[0], cell_data_to_object));
            list.AddRange(objects);
            return list;
        }

        protected static IList<T> _GetCells<T, TResult>(
            IVisio.Shape shape,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], T> cell_data_to_object)
        {
            CellGroupMultiRow.verify_multirow_query(query);

            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetFormulasAndResults<TResult>(surface);
            var sec = data_for_shape.Sections[0];
            var sec_objects = CellGroupMultiRow.SectionRowsToObjects(sec, cell_data_to_object);
            
            return sec_objects;
        }

        private static List<T> SectionRowsToObjects<T, TResult>(SubQueryOutput<ShapeSheet.CellData<TResult>> sec, System.Func<ShapeSheet.CellData<TResult>[],T> cells_to_object)
        {
            var sec_objects = new List<T>(sec.Rows.Count);
            var objects = sec.Rows.Select(row => cells_to_object(row.Cells));
            sec_objects.AddRange(objects);
            return sec_objects;
        }

        public void SetFormulas(short shapeid, FormulaWriterSIDSRC writer, short row)
        {
            foreach (var pair in this.Pairs)
            {
                var new_src = pair.SRC.CopyWithNewRow(row);
                writer.SetFormula(shapeid, new_src, pair.Formula);
            }
        }

        public void SetFormulas(FormulaWriterSRC writer, short row)
        {
            foreach (var pair in this.Pairs)
            {
                var new_src = pair.SRC.CopyWithNewRow(row);
                writer.SetFormula(new_src, pair.Formula);
            }
        }

    }
}
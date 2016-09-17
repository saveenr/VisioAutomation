using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Queries;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRow : CellGroupBase
    {
        private static void verify_singlerow_query(Query query)
        {
            if (query.Cells.Count < 1)
            {
                throw new InternalAssertionException("Query must contain at least one cell");
            }

            if (query.SubQueries.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain contain any subqueries");
            }
        }

        protected static List<TCellGroup> _GetCells<TCellGroup, TResult>(
            IVisio.Page page, 
            IList<int> shapeids,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], TCellGroup> cells_to_object)
        {
            verify_singlerow_query(query);

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<TCellGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => cells_to_object(d.Cells));
            list.AddRange(objects);
            return list;
        }

        protected static TCellGroup _GetCells<TCellGroup, TResult>(
            IVisio.Shape shape,
            Query query,
            System.Func<ShapeSheet.CellData<TResult>[], TCellGroup> cells_to_object)
        {
            verify_singlerow_query(query);

            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = query.GetFormulasAndResults<TResult>(surface);
            var cells = cells_to_object(data_for_shape.Cells);
            return cells;
        }

        public void SetFormulas(FormulaWriterSRC writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(pair.SRC, pair.Formula);
            }
        }

        public void SetFormulas(short shapeid, FormulaWriterSIDSRC writer)
        {
            foreach (var pair in this.Pairs)
            {
                writer.SetFormula(shapeid, pair.SRC, pair.Formula);
            }
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    public abstract class CellGroupSingleRowQuery<TCellGroup,TResult>: CellGroupQuery<TCellGroup, TResult>
    {

        protected override void validate_query()
        {
            if (this.query.Cells.Count < 1)
            {
                throw new InternalAssertionException("Query must contain at least one cell");
            }

            if (this.query.SubQueries.Count != 0)
            {
                throw new InternalAssertionException("Query should not contain contain any subqueries");
            }
        }

        public List<TCellGroup> GetCells(
            Microsoft.Office.Interop.Visio.Page page,
            IList<int> shapeids)
        {
            validate_query();

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = this.query.GetFormulasAndResults<TResult>(surface, shapeids);
            var list = new List<TCellGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TCellGroup GetCells(
            Microsoft.Office.Interop.Visio.Shape shape)
        {
            validate_query();
            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = this.query.GetFormulasAndResults<TResult>(surface);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
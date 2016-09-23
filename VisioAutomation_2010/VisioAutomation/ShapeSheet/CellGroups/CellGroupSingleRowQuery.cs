using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupSingleRowQuery<TCellGroup>: CellGroupQuery<TCellGroup>
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

        public List<TCellGroup> GetCellGroups(IVisio.Page page, IList<int> shapeids)
        {
            validate_query();

            var surface = new ShapeSheetSurface(page);
            var data_for_shapes = this.query.GetFormulasAndResults(surface, shapeids);
            var list = new List<TCellGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TCellGroup GetCellGroup(IVisio.Shape shape)
        {
            validate_query();
            var surface = new ShapeSheetSurface(shape);
            var data_for_shape = this.query.GetFormulasAndResults(surface);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
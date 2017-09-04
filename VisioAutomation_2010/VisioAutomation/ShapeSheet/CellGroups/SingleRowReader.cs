using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class SingleRowReader<TCellGroup>
    {
        protected Query.CellQuery query;

        protected SingleRowReader()
        {
            this.query = new Query.CellQuery();
        }

        public abstract TCellGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

        protected void validate_query()
        {
            if (this.query.Cells.Count < 1)
            {
                throw new InternalAssertionException("Query must contain at least one cell");
            }
        }

        public List<TCellGroup> GetCellGroups(IVisio.Page page, IList<int> shapeids)
        {
            validate_query();

            var data_for_shapes = this.query.GetFormulasAndResults(page, shapeids);
            var list = new List<TCellGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TCellGroup GetCellGroup(IVisio.Shape shape)
        {
            validate_query();
            var data_for_shape = this.query.GetFormulasAndResults(shape);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
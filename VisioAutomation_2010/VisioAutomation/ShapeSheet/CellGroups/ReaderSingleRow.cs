using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderSingleRow<TGroup> where TGroup : CellGroupSingleRow
    {
        protected Query.CellQuery query;

        protected ReaderSingleRow()
        {
            this.query = new Query.CellQuery();
        }

        public abstract TGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

        public List<TGroup> GetCellGroups(IVisio.Page page, IList<int> shapeids)
        {
            var data_for_shapes = this.query.GetFormulasAndResults(page, shapeids);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetCellGroup(IVisio.Shape shape)
        {
            var data_for_shape = this.query.GetFormulasAndResults(shape);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
using System.Collections.Generic;
using System.Linq;
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

        public abstract TGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row);

        public List<TGroup> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var data_for_shapes = this.query.GetCells(page, shapeids, cvt);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var data_for_shape = this.query.GetCells(shape, cvt);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
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

        public List<TGroup> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var data_for_shapes = this.query.GetFormulas(page, shapeids);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public List<TGroup> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var data_for_shapes = this.query.GetResults<string>(page, shapeids);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetFormulas(IVisio.Shape shape)
        {
            var data_for_shape = this.query.GetFormulas(shape);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }

        public TGroup GetResults(IVisio.Shape shape)
        {
            var data_for_shape = this.query.GetResults<string>(shape);
            var cells = this.CellDataToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
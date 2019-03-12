using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderSingleRow<TGroup> where TGroup : CellGroupSingleRow
    {
        protected Query.CellQuery query;

        protected ReaderSingleRow()
        {
            this.query = new VASS.Query.CellQuery();
        }

        public abstract TGroup ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row);

        public List<TGroup> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var data_for_shapes = this.query.GetCells(page, shapeids, type);
            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.ToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetCells(IVisio.Shape shape, CellValueType type)
        {
            var data_for_shape = this.query.GetCells(shape, type);
            var cells = this.ToCellGroup(data_for_shape.Cells);
            return cells;
        }
    }
}
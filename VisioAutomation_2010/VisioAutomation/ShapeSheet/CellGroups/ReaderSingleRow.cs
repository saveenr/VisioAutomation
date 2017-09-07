using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.ShapeSheet
{
    public enum CellValueType
    {
        Formula, Result
    }
}

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

        public List<TGroup> GetCellGroups(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            CellQueryOutputList<string> data_for_shapes;

            if (cvt == CellValueType.Formula)
            {
                data_for_shapes = this.query.GetFormulas(page, shapeids);
            }
            else
            {
                data_for_shapes = this.query.GetResults<string>(page, shapeids);

            }

            var list = new List<TGroup>(shapeids.Count);
            var objects = data_for_shapes.Select(d => this.CellDataToCellGroup(d.Cells));
            list.AddRange(objects);
            return list;
        }

        public TGroup GetCellGroup(IVisio.Shape shape, CellValueType cvt)
        {
            if (cvt == CellValueType.Formula)
            {
                var data_for_shape = this.query.GetFormulas(shape);
                var cells = this.CellDataToCellGroup(data_for_shape.Cells);
                return cells;
            }
            else
            {
                var data_for_shape = this.query.GetResults<string>(shape);
                var cells = this.CellDataToCellGroup(data_for_shape.Cells);
                return cells;

            }
        }
    }
}
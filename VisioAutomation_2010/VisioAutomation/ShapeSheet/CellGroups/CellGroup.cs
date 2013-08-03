using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        public abstract void ApplyFormulas(ApplyFormula func);

        protected static IList<T> _GetCells<T>(IVisio.Page page, IList<int> shapeids, VA.ShapeSheet.Query.CellQuery cellQuery, QueryResultToObject<T> f)
        {
            var data_for_shapes = cellQuery.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<T>(shapeids.Count);
            foreach (var data_for_shape in data_for_shapes)
            {
                var cells = f(data_for_shape);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T>(IVisio.Shape shape, VA.ShapeSheet.Query.CellQuery cellQuery, QueryResultToObject<T> f)
        {
            var data_for_shape = cellQuery.GetFormulasAndResults<double>(shape);
            var cells = f(data_for_shape);
            return cells;
        }
    }
}
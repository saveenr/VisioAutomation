using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroup : BaseCellGroup
    {
        public abstract void ApplyFormulas(ApplyFormula func);

        protected static IList<T> _GetCells<T>(IVisio.Page page, IList<int> shapeids, CellQuery cellQuery, ResultToCells<T> f)
        {
            var data = cellQuery.GetFormulasAndResults<double>(page, shapeids);
            var list = new List<T>();
            for (int i = 0; i < shapeids.Count; i++)
            {
                var cells = f(data[i]);
                list.Add(cells);
            }
            return list;
        }

        protected static T _GetCells<T>(IVisio.Shape shape, CellQuery cellQuery, ResultToCells<T> f)
        {
            var data_for_shape = cellQuery.GetFormulasAndResults<double>(shape);
            var cells = f(data_for_shape);
            return cells;
        }
    }
}
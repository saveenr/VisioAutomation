using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        // This class is meant for those cell groups that appear as multiple rows in a section
        // for example the character section or the paragraph section

        public abstract void ApplyFormulasForRow(ApplyFormula func, short row);

        public static IList<List<T>> _GetCells<T>(IVisio.Page page, IList<int> shapeids, VA.ShapeSheet.Query.CellQuery cellQuery, RowToObject<T> f)
        {
            var outer_list = new List<List<T>>();
            var data_for_shapes = cellQuery.GetFormulasAndResults<double>(page, shapeids);

            foreach (var data_for_shape in data_for_shapes)
            {
                var inner_list = new List<T>();
                outer_list.Add(inner_list);

                var sec = data_for_shape.SectionCells[0];
                foreach (var row in sec)
                {
                    var obj = f(row);
                    inner_list.Add(obj);
                }
            }
            return outer_list;
        }

        public static IList<T> _GetCells<T>(IVisio.Shape shape, VA.ShapeSheet.Query.CellQuery cellQuery, RowToObject<T> f)
        {
            var data_for_shape = cellQuery.GetFormulasAndResults<double>(shape);
            var inner_list = new List<T>();
            var sec = data_for_shape.SectionCells[0];
            foreach (var row in sec)
            {
                var obj = f(row);
                inner_list.Add(obj);
            }
            return inner_list;
        }
    }
}
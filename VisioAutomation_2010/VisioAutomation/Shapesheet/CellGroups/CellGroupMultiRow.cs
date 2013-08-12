using System.Security.Authentication;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        // This class is meant for those cell groups that appear as multiple rows in a section
        // for example the character section or the paragraph section

        // Note: in the _GetCells method below you will notice that SectionCells[0] is used
        // Why 0? In the context of how these methods are used, only one section is retrieved so
        // that's why ony the first section result is retrieved - there are no more.

        public abstract void ApplyFormulasForRow(ApplyFormula func, short row);

        public static IList<List<T>> _GetCells<T>(
            IVisio.Page page, IList<int> shapeids, 
            VA.ShapeSheet.Query.CellQuery cellQuery, 
            RowToObject<T> f)
        {
            var outer_list = new List<List<T>>();
            var data_for_shapes = cellQuery.GetFormulasAndResults<double>(page, shapeids);


            foreach (var data_for_shape in data_for_shapes)
            {
                if (data_for_shape.SectionCells.Count != 1)
                {
                    var msg = string.Format("Internal Error: Only 1 section should be in these queries");
                    throw new AuthenticationException(msg);
                }

                var sec = data_for_shape.SectionCells[0];
                var inner_list = new List<T>(sec.Count);

                foreach (var row in sec)
                {
                    var obj = f(row);
                    inner_list.Add(obj);
                }

                outer_list.Add(inner_list);
            }
            return outer_list;
        }

        public static IList<T> _GetCells<T>(
            IVisio.Shape shape, 
            VA.ShapeSheet.Query.CellQuery cellQuery, 
            RowToObject<T> f)
        {
            var data_for_shape = cellQuery.GetFormulasAndResults<double>(shape);

            if (data_for_shape.SectionCells.Count != 1)
            {
                var msg = string.Format("Internal Error: Only 1 section should be in these queries");
                throw new AuthenticationException(msg);
            }

            var sec = data_for_shape.SectionCells[0];
            var inner_list = new List<T>(sec.Count);
            
            foreach (var row in sec)
            {
                var obj = f(row);
                inner_list.Add(obj);
            }
            
            return inner_list;
        }
    }
}
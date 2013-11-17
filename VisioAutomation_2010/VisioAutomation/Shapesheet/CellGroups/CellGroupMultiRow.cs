using System.Security.Authentication;
using System.Xml.Serialization;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class CellGroupMultiRow : BaseCellGroup
    {
        private static void check_query(VA.ShapeSheet.Query.CellQuery query)
        {
            if (query.Columns.Count != 0)
            {
                throw new VA.AutomationException("Query should not contain any Columns");
            }

            if (query.Sections.Count != 1)
            {
                throw new VA.AutomationException("Query should not contain contain exaxtly 1 section");
            }            
        }

        public static IList<List<T>> _GetCells<T>(
            IVisio.Page page, 
            IList<int> shapeids, 
            VA.ShapeSheet.Query.CellQuery query, 
            RowToObject<T> f)
        {
            check_query(query);

            var outer_list = new List<List<T>>();
            var data_for_shapes = query.GetFormulasAndResults<double>(page, shapeids);

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
            VA.ShapeSheet.Query.CellQuery query, 
            RowToObject<T> f)
        {
            check_query(query);

            var data_for_shape = query.GetFormulasAndResults<double>(shape);

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
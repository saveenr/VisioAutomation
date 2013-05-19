using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS
{
    static class VisioPSUtil
    {
        public static System.Data.DataTable todatatable<T>(VA.ShapeSheet.Data.Table<T> output, IList<string> names)
        {
            var dt = new System.Data.DataTable();
            foreach (string name in names)
            {
                dt.Columns.Add(name, typeof(T));
            }
            int colcount = names.Count;
            var arr = new object[colcount];
            for (int r = 0; r < output.RowCount; r++)
            {
                for (int i = 0; i < colcount; i++)
                {
                    arr[i] = output[r, i];
                }
                dt.Rows.Add(arr);
            }
            return dt;
        }

        public static System.Data.DataTable QueryToDataTable(VA.ShapeSheet.Query.CellQuery query, bool getresults, ResultType ResultType, IList<int> shapeids, IVisio.Page page)
        {
            var names = query.Columns.Select(c => c.Name).ToList();
            if (getresults)
            {
                if (ResultType == ResultType.String)
                {
                    var output = query.GetResults<string>(page, shapeids);
                    return VisioPSUtil.todatatable(output, names);
                }
                else if (ResultType == ResultType.Boolean)
                {
                    var output = query.GetResults<bool>(page, shapeids);
                    return VisioPSUtil.todatatable(output, names);
                }
                else if (ResultType == ResultType.Double)
                {
                    var output = query.GetResults<double>(page, shapeids);
                    return VisioPSUtil.todatatable(output, names);
                }
                else if (ResultType == ResultType.Integer)
                {
                    var output = query.GetResults<int>(page, shapeids);
                    return VisioPSUtil.todatatable(output, names);
                }
                else
                {
                    throw new VA.Scripting.VisioApplicationException("Unsupported Result type");
                }

            }
            else
            {
                var output = query.GetFormulas(page, shapeids);
                return VisioPSUtil.todatatable(output, names);
            }
        }

    }
}

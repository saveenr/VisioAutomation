using System.Collections.Generic;
using System.Data;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS
{
    static class VisioPSUtil
    {
        public static DataTable querytable_to_datatable<T>(CellQuery cellQuery, CellQuery.QueryResultList<T> query_output)
        {
            // First Construct a Datatable with a compatible schema
            var dt = new System.Data.DataTable();
            foreach (var col in cellQuery.Columns)
            {
                dt.Columns.Add(col.Name, typeof(T));
            }

            // Then populate the rows of the datatable
            dt.BeginLoadData();
            int colcount = cellQuery.Columns.Count;
            var rowbuf = new object[colcount];
            for (int r = 0; r < query_output.Count; r++)
            {
                // populate the row buffer
                for (int i = 0; i < colcount; i++)
                {
                    rowbuf[i] = query_output[r][i];
                }

                // load it into the table
                dt.Rows.Add(rowbuf);
            }
            dt.EndLoadData();
            return dt;
        }

        public static System.Data.DataTable QueryToDataTable(VA.ShapeSheet.Query.CellQuery cellQuery, bool getresults, ResultType ResultType, IList<int> shapeids, IVisio.Page page)
        {
            if (getresults)
            {
                if (ResultType == ResultType.String)
                {
                    var output = cellQuery.GetResults<string>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Boolean)
                {
                    var output = cellQuery.GetResults<bool>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Double)
                {
                    var output = cellQuery.GetResults<double>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Integer)
                {
                    var output = cellQuery.GetResults<int>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(cellQuery, output);
                }
                else
                {
                    throw new VA.Scripting.VisioApplicationException("Unsupported Result type");
                }
            }
            else
            {
                var output = cellQuery.GetFormulas(page, shapeids);
                return VisioPSUtil.querytable_to_datatable(cellQuery, output);
            }
        }
    }
}

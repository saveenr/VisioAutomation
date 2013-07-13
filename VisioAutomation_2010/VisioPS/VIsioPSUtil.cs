using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using VisioAutomation.ShapeSheet.Data;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS
{
    static class VisioPSUtil
    {
        public static DataTable querytable_to_datatable<T>(QueryEx query, List<ExQueryResult<T>> query_output)
        {
            // TODO Add Name
            // First Construct a Datatable with a compatible schema
            var dt = new System.Data.DataTable();
            foreach (var col in query.Columns)
            {
                dt.Columns.Add("CELLNAME", typeof(T));
            }

            // Then populate the rows of the datatable
            dt.BeginLoadData();
            int colcount = query.Columns.Count;
            var rowbuf = new object[colcount];
            for (int r = 0; r < query_output.Count; r++)
            {
                // populate the row buffer
                for (int i = 0; i < colcount; i++)
                {
                    rowbuf[i] = query_output[r].Cells[i];
                }

                // load it into the table
                dt.Rows.Add(rowbuf);
            }
            dt.EndLoadData();
            return dt;
        }

        public static System.Data.DataTable QueryToDataTable(VA.ShapeSheet.Query.QueryEx query, bool getresults, ResultType ResultType, IList<int> shapeids, IVisio.Page page)
        {
            if (getresults)
            {
                if (ResultType == ResultType.String)
                {
                    var output = query.GetResults<string>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(query, output);
                }
                else if (ResultType == ResultType.Boolean)
                {
                    var output = query.GetResults<bool>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(query, output);
                }
                else if (ResultType == ResultType.Double)
                {
                    var output = query.GetResults<double>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(query, output);
                }
                else if (ResultType == ResultType.Integer)
                {
                    var output = query.GetResults<int>(page, shapeids);
                    return VisioPSUtil.querytable_to_datatable(query, output);
                }
                else
                {
                    throw new VA.Scripting.VisioApplicationException("Unsupported Result type");
                }
            }
            else
            {
                var output = query.GetFormulas(page, shapeids);
                return VisioPSUtil.querytable_to_datatable(query, output);
            }
        }
    }
}

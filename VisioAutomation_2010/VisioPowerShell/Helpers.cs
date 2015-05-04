using System.Collections.Generic;
using System.Data;
using VisioAutomation.Scripting;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell
{
    static class Helpers
    {
        private static DataTable querytable_to_datatable<T>(CellQuery cellQuery, QueryResultList<T> query_output)
        {
            // First Construct a Datatable with a compatible schema
            var dt = new DataTable();
            dt.Columns.Add("ShapeID", typeof(int));
            foreach (var col in cellQuery.CellColumns)
            {
                dt.Columns.Add(col.Name, typeof(T));
            }

            // Then populate the rows of the datatable
            dt.BeginLoadData();
            int colcount = cellQuery.CellColumns.Count;
            var rowbuf = new object[colcount+1];
            for (int r = 0; r < query_output.Count; r++)
            {
                // populate the row buffer

                rowbuf[0] = query_output[r].ShapeID;

                for (int i = 0; i < colcount; i++)
                {
                    rowbuf[i+1] = query_output[r][i];
                }

                // load it into the table
                dt.Rows.Add(rowbuf);
            }
            dt.EndLoadData();
            return dt;
        }

        public static DataTable QueryToDataTable(CellQuery cellQuery, bool getresults, ResultType ResultType, IList<int> shapeids, ShapeSheetSurface surface)
        {
            if (getresults)
            {
                if (ResultType == ResultType.String)
                {
                    var output = cellQuery.GetResults<string>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Boolean)
                {
                    var output = cellQuery.GetResults<bool>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Double)
                {
                    var output = cellQuery.GetResults<double>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                else if (ResultType == ResultType.Integer)
                {
                    var output = cellQuery.GetResults<int>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                else
                {
                    throw new VisioApplicationException("Unsupported Result type");
                }
            }
            else
            {
                var output = cellQuery.GetFormulas(surface, shapeids);
                return Helpers.querytable_to_datatable(cellQuery, output);
            }
        }
    }
}

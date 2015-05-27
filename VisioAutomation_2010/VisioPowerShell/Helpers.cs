using System.Collections.Generic;
using System.Data;

namespace VisioPowerShell
{
    static class Helpers
    {
        private static DataTable querytable_to_datatable<T>(VisioAutomation.ShapeSheet.Query.CellQuery cellQuery, VisioAutomation.ShapeSheet.Query.QueryResultList<T> query_output)
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

        public static DataTable QueryToDataTable(VisioAutomation.ShapeSheet.Query.CellQuery cellQuery, bool getresults, Model.ResultType ResultType, IList<int> shapeids, VisioAutomation.ShapeSheet.ShapeSheetSurface surface)
        {
            if (!getresults)
            {
                var output = cellQuery.GetFormulas(surface, shapeids);
                return Helpers.querytable_to_datatable(cellQuery, output);
            }

            switch (ResultType)
            {
                case Model.ResultType.String:
                {
                    var output = cellQuery.GetResults<string>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                case Model.ResultType.Boolean:
                {
                    var output = cellQuery.GetResults<bool>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                case Model.ResultType.Double:
                {
                    var output = cellQuery.GetResults<double>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
                case Model.ResultType.Integer:
                {
                    var output = cellQuery.GetResults<int>(surface, shapeids);
                    return Helpers.querytable_to_datatable(cellQuery, output);
                }
            }

            throw new VisioAutomation.Scripting.VisioApplicationException("Unsupported Result type");
        }
    }
}

using System.Collections.Generic;
using System.Data;
using VisioAutomation.ShapeSheet.Query;

namespace VisioPowerShell.Models
{
    static class DataTableHelpers
    {
        private static DataTable querytable_to_datatable<T>(
            CellQuery cell_query, 
            CellQueryOutputList<T> query_output)
        {
            // First Construct a Datatable with a compatible schema
            var dt = new DataTable();
            foreach (var col in cell_query.Columns)
            {
                dt.Columns.Add(col.Name, typeof(T));
            }

            // Then populate the rows of the datatable
            dt.BeginLoadData();
            int colcount = cell_query.Columns.Count;
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

        public static DataTable QueryToDataTable(CellQuery cell_query,CellOutputType cell_output_type, IList<int> shapeids, VisioAutomation.SurfaceTarget surface)
        {
            switch (cell_output_type)
            {
                case CellOutputType.Formula:
                {
                    var output = cell_query.GetFormulas(surface, shapeids);
                    var dt = DataTableHelpers.querytable_to_datatable(cell_query, output);
                    return dt;
                }
                case CellOutputType.ResultString:
                {
                    var output = cell_query.GetResults<string>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(cell_query, output);
                }
                case CellOutputType.ResultBoolean:
                {
                    var output = cell_query.GetResults<bool>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(cell_query, output);
                }
                case CellOutputType.ResultDouble:
                {
                    var output = cell_query.GetResults<double>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(cell_query, output);
                }
                case CellOutputType.ResultInteger:
                {
                    var output = cell_query.GetResults<int>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(cell_query, output);
                }
            }

            string msg = string.Format("Unsupported value of \"{0}\" for type {1}", cell_output_type, nameof(CellOutputType));
            throw new System.ArgumentOutOfRangeException(nameof(cell_output_type), msg);
        }
    }
}

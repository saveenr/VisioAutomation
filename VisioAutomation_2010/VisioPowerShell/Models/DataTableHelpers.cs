using System.Collections.Generic;
using System.Data;
using VisioAutomation.ShapeSheet.Query;

namespace VisioPowerShell.Models
{
    static class DataTableHelpers
    {
        private static DataTable querytable_to_datatable<T>(
            CellQuery query, 
            Rows<T> output)
        {
            // First Construct a Datatable with a compatible schema
            var dt = new DataTable();
            foreach (var col in query.Columns)
            {
                dt.Columns.Add(col.Name, typeof(T));
            }

            // Then populate the rows of the datatable
            dt.BeginLoadData();
            int colcount = query.Columns.Count;
            var rowbuf = new object[colcount];
            for (int row_index = 0; row_index < output.Count; row_index++)
            {
                // populate the row buffer
                for (int col_index = 0; col_index < colcount; col_index++)
                {
                    rowbuf[col_index] = output[row_index][col_index];
                }

                // load it into the table
                dt.Rows.Add(rowbuf);
            }
            dt.EndLoadData();
            return dt;
        }

        public static DataTable QueryToDataTable(
            CellQuery query,
            CellOutputType output_type,
            IList<int> shapeids, 
            VisioAutomation.SurfaceTarget surface)
        {
            switch (output_type)
            {
                case CellOutputType.Formula:
                {
                    var output = query.GetFormulas(surface, shapeids);
                    var dt = DataTableHelpers.querytable_to_datatable(query, output);
                    return dt;
                }
                case CellOutputType.ResultString:
                {
                    var output = query.GetResults<string>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(query, output);
                }
                case CellOutputType.ResultBoolean:
                {
                    var output = query.GetResults<bool>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(query, output);
                }
                case CellOutputType.ResultDouble:
                {
                    var output = query.GetResults<double>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(query, output);
                }
                case CellOutputType.ResultInteger:
                {
                    var output = query.GetResults<int>(surface, shapeids);
                    return DataTableHelpers.querytable_to_datatable(query, output);
                }
            }

            string msg = string.Format("Unsupported value of \"{0}\" for type {1}", output_type, nameof(CellOutputType));
            throw new System.ArgumentOutOfRangeException(nameof(output_type), msg);
        }
    }
}

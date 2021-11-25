using System.Data;
using VisioPowerShell.Models;


namespace VisioPowerShell.Internal;

static class DataTableHelpers
{
    private static DataTable querytable_to_datatable<T>(
        VASS.Query.CellQuery query,
        VASS.Query.Rows<T> output)
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
        VASS.Query.CellQuery query,
        VisioAutomation.ShapeSheet.CellValueType value_type,
        Models.ResultType result_type,
        IList<int> shapeids, 
        VisioAutomation.SurfaceTarget surface)
    {

        if (value_type == VASS.CellValueType.Formula)
        {
            var output = query.GetFormulas(surface, shapeids);
            var dt = DataTableHelpers.querytable_to_datatable(query, output);
            return dt;
        }

        if (value_type != VASS.CellValueType.Result)
        {
            throw new System.ArgumentOutOfRangeException(nameof(value_type));
        }

        if (result_type == ResultType.String)
        {
            var output = query.GetResults<string>(surface, shapeids);
            return DataTableHelpers.querytable_to_datatable(query, output);
        }
        else if (result_type == ResultType.Bool)
        {
            var output = query.GetResults<string>(surface, shapeids);
            return DataTableHelpers.querytable_to_datatable(query, output);
        }
        else if (result_type == ResultType.Double)
        {
            var output = query.GetResults<double>(surface, shapeids);
            return DataTableHelpers.querytable_to_datatable(query, output);
        }
        else if(result_type == ResultType.Int)
        {
            var output = query.GetResults<int>(surface, shapeids);
            return DataTableHelpers.querytable_to_datatable(query, output);
        }
        else
        {
            string msg = string.Format("Unsupported value of \"{0}\" for type {1}", result_type, nameof(result_type));
            throw new System.ArgumentOutOfRangeException(nameof(result_type), msg);
        }

    }
}
using VA=VisioAutomation;

namespace VisioPS
{
    internal class DTUtil
    {
        public static System.Data.DataTable ToDataTable<T>(VA.ShapeSheet.Data.Table<T> table)
        {
            int extra_columns = 1;
            var datatable = new System.Data.DataTable();

            // define the extra columns
            datatable.Columns.Add("ShapeID", typeof(int));

            // define the normal columns
            for (int i = 0; i < table.Columns.Count; i++)
            {
                var colname = string.Format("Cell{0}", i);
                datatable.Columns.Add(colname, typeof(T));
            }

            var rowarray = new object[table.Columns.Count + extra_columns];
            
            // Fill in the rows
            datatable.BeginLoadData();
            foreach (var group in table.Groups.items)
            {
                foreach (int i in group.RowIndices)
                {
                    // set values for the extra columns
                    rowarray[0] = group.ShapeID;

                    // set values for the cell columns
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        rowarray[c + extra_columns] = table[i, c];
                    }
                    datatable.Rows.Add(rowarray);
                }
            }
            datatable.EndLoadData();

            return datatable;
        }
    }
}
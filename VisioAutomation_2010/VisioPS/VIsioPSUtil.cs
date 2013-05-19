using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA = VisioAutomation;

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
    }
}

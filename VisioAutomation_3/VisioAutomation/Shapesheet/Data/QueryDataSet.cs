using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
namespace VisioAutomation.ShapeSheet.Data
{
    public class QueryDataSet<T>
    {
        internal readonly int RowCount;
        internal readonly int ColumnCount;

        public TableRowGroupList Groups { get; private set; }
        public Table<string> Formulas { get; private set; }
        public Table<T> Results { get; private set; }

        internal QueryDataSet(string[] formulas_array, T[] results_array, IList<int> shapeids, int columncount,
                            int rowcount, IList<int> groupcounts)
        {
            if (formulas_array == null && results_array == null)
            {
                throw new AutomationException("Both formulas and results cannot be null");
            }

            if (formulas_array != null & results_array != null)
            {
                if (formulas_array.Length != results_array.Length)
                {
                    throw new AutomationException("Formula array and Result array must have the same length");
                }
            }

            if (shapeids.Count != groupcounts.Count)
            {
                throw new AutomationException("The number of shapes must be equal to the number of groups");
            }

            int groupcountsum = groupcounts.Sum();
            if (rowcount != groupcountsum)
            {
                throw new AutomationException("The total number of rows must be equal to the sum of the group counts");                
            }

            int totalcells = columncount*rowcount;

            if (formulas_array != null)
            {
                if (totalcells != formulas_array.Length)
                {
                    throw new AutomationException("Invalid number of formulas");
                }                
            }

            if (results_array != null)
            {
                if (totalcells != results_array.Length)
                {
                    throw new AutomationException("Invalid number of formulas");
                }
            }

            this.RowCount = rowcount;
            this.ColumnCount = columncount;

            this.Groups = new TableRowGroupList();
            foreach (var g in this.GetGrouping(shapeids, groupcounts))
            {
                this.Groups.Add(g);
            }

            this.Formulas = formulas_array != null ? this.FromDataSet<string>(i => formulas_array[i]) : null;
            this.Results = results_array != null ? this.FromDataSet<T>(i => results_array[i]) : null;
        }

        private TableRowGroup[] GetGrouping(IList<int> shape_ids, IList<int> group_counts)
        {
            var table_total_rows = this.RowCount;

            if (group_counts == null)
            {
                throw new System.ArgumentNullException("group_counts");
            }

            if (group_counts.Count != shape_ids.Count)
            {
                string msg = String.Format("Number of group counts {0} does not match number of shape ids {1}",
                                           group_counts.Count, shape_ids.Count);
                throw new AutomationException(msg);
            }

            // Group the rows
            var groups = new TableRowGroup[group_counts.Count];
            int cur_group_start = 0;

            for (int i = 0; i < group_counts.Count; i++)
            {
                int cur_group_count = group_counts[i];
                var cur_group_shape_id = shape_ids[i];

                if (cur_group_count < 1)
                {
                    // the group has no rows, so create an empty RowGroup
                    var new_group = new TableRowGroup(cur_group_shape_id, cur_group_count, -1, -1);
                    groups[i] = new_group;
                }
                else
                {
                    // the group contains at least 1 row create a non-empty RowGroup
                    int new_group_start = cur_group_start;
                    int new_group_end = new_group_start + cur_group_count - 1;
                    var new_group = new TableRowGroup(cur_group_shape_id, cur_group_count, new_group_start, new_group_end);
                    groups[i] = new_group;

                    // update the new starting position for the next group
                    cur_group_start += cur_group_count;
                }
            }

            // verify that the groups account for all the rows
            var total_rows_in_groups = groups.Select(g => g.Count).Sum();
            if (total_rows_in_groups != table_total_rows)
            {
                throw new AutomationException("Internal Error: rows in groups and total rows do not match");
            }

            return groups;
        }

        private Table<X> FromDataSet<X>(System.Func<int, X> get_data_at_position)
        {
            var vals = new X[this.RowCount,this.ColumnCount];
            for (int r = 0; r < this.RowCount; r++)
            {
                for (int c = 0; c < this.ColumnCount; c++)
                {
                    int i = (r * this.ColumnCount) + c;
                    vals[r, c] = get_data_at_position(i);
                }
            }

            var table = new Table<X>(this.RowCount, this.ColumnCount, this.Groups, vals);
            return table;
        }

        public QueryDataRow<T> GetRow(int row)
        {
            return new QueryDataRow<T>(this, row);
        }

        public VA.ShapeSheet.CellData<T> GetItem(int row, VA.ShapeSheet.Query.QueryColumn col)
        {
            string formula = this.Formulas[row, col];
            T result = this.Results[row, col];
            var cd = new VA.ShapeSheet.CellData<T>(formula, result);
            return cd;
        }

        internal IEnumerable<QueryDataRow<T>> EnumRows()
        {
            for (int row = 0; row < this.RowCount; row++)
            {
                yield return this.GetRow(row);
            }
        }
    }
}
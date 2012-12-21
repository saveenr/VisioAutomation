using System.Collections.Generic;
using System.Collections;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.ShapeSheet.Data
{
    public class TableRowGroupList : IEnumerable<TableRowGroup>
    {
        public List<TableRowGroup> items;

        internal static TableRowGroupList Build(IList<int> shapeids, IList<int> groupcounts, int rowcount)
        {
            var groups = new TableRowGroupList();
            foreach (var g in GetGrouping(shapeids, groupcounts, rowcount))
            {
                groups.Add(g);
            }
            return groups;
        }

        internal static VA.ShapeSheet.Data.TableRowGroup[] GetGrouping(IList<int> shape_ids, IList<int> group_counts, int table_total_rows)
        {
            if (group_counts == null)
            {
                throw new System.ArgumentNullException("group_counts");
            }

            if (group_counts.Count != shape_ids.Count)
            {
                string msg = string.Format("Number of group counts {0} does not match number of shape ids {1}",
                                           group_counts.Count, shape_ids.Count);
                throw new AutomationException(msg);
            }

            // Group the rows
            var groups = new VA.ShapeSheet.Data.TableRowGroup[group_counts.Count];
            int cur_group_start = 0;

            for (int i = 0; i < group_counts.Count; i++)
            {
                int cur_group_count = group_counts[i];
                var cur_group_shape_id = shape_ids[i];

                if (cur_group_count < 1)
                {
                    // the group has no rows, so create an empty RowGroup
                    var new_group = new VA.ShapeSheet.Data.TableRowGroup(cur_group_shape_id, cur_group_count, -1, -1);
                    groups[i] = new_group;
                }
                else
                {
                    // the group contains at least 1 row create a non-empty RowGroup
                    int new_group_start = cur_group_start;
                    int new_group_end = new_group_start + cur_group_count - 1;
                    var new_group = new VA.ShapeSheet.Data.TableRowGroup(cur_group_shape_id, cur_group_count, new_group_start, new_group_end);
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


        public TableRowGroupList()
        {
            this.items = new List<TableRowGroup>();
        }

        public void Add(TableRowGroup g)
        {
            this.items.Add(g);
        }

        public TableRowGroup this[int index]
        {
            get { return this.items[index]; }
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public IEnumerator<TableRowGroup> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return GetEnumerator();
        }
    }
}
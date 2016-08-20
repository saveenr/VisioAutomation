using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Columns
{
    public class ListColumnSubQuery : ListColumnBase<ColumnSubQuery>
    {
        private HashSet<short> _cellindex_set;

        internal ListColumnSubQuery() :
            base(0)
        {
        }

        internal ListColumnSubQuery(int capacity) : base(capacity)
        {
        }

        public ColumnSubQuery Add(short cellindex, string sname)
        {
            check_duplicate_cellindex(cellindex);
            string norm_name = this.normalize_name(sname);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new ColumnSubQuery(ordinal, cellindex, norm_name);
            this._items.Add(col);
            this._cellindex_set.Add(cellindex);

            return col;
        }

        private void check_duplicate_cellindex(short cellindex)
        {
            if (this._cellindex_set == null)
            {
                this._cellindex_set = new HashSet<short>();
            }

            if (this._cellindex_set.Contains(cellindex))
            {
                string msg = string.Format("Duplicate Cell Index: {0}", cellindex);
                throw new AutomationException(msg);
            }
        }
    }
}
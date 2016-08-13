using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Columns
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

        public ColumnSubQuery Add(short cell, string name)
        {
            if (this._cellindex_set == null)
            {
                this._cellindex_set = new HashSet<short>();
            }

            if (this._cellindex_set.Contains(cell))
            {
                string msg = "Duplicate Cell Index";
                throw new AutomationException(msg);
            }

            name = this.fixup_name(name);
            int ordinal = this._items.Count;
            var col = new ColumnSubQuery(ordinal, cell, name);
            this._items.Add(col);
            this._cellindex_set.Add(cell);
            return col;
        }
    }
}
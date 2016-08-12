using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery.Columns
{
    public class ListColumnCellIndex : ListColumnBase<ColumnCellIndex>
    {
        private HashSet<short> _cellindex_set;

        internal ListColumnCellIndex() :
            base(0)
        {
        }

        internal ListColumnCellIndex(int capacity) : base(capacity)
        {
        }

        public ColumnCellIndex Add(short cell, string name)
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
            var col = new ColumnCellIndex(ordinal, cell, name);
            this._items.Add(col);
            this._cellindex_set.Add(cell);
            return col;
        }
    }
}
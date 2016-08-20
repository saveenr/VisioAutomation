using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Columns
{
    public class ListColumnQuery : ListColumnBase<ColumnQuery>
    {
        private HashSet<ShapeSheet.SRC> _src_set;

        internal ListColumnQuery() :
            this(0)
        {
        }

        internal ListColumnQuery(int capacity) : base(capacity)
        {
        }

        internal ColumnQuery Add(ShapeSheet.SRC src) => this.Add(src, null);

        internal ColumnQuery Add(ShapeSheet.SRC src, string name)
        {
            name = this.fixup_name(name);

            if (this._dic_columns.ContainsKey(name))
            {
                throw new AutomationException("Duplicate Column Name");
            }

            if (this._src_set == null)
            {
                this._src_set = new HashSet<ShapeSheet.SRC>();
            }

            if (this._src_set.Contains(src))
            {
                string msg = "Duplicate SRC";
                throw new AutomationException(msg);
            }

            int ordinal = this._items.Count;
            var col = new ColumnQuery(ordinal, src, name);
            this._items.Add(col);

            this._dic_columns[name] = col;
            this._src_set.Add(src);
            return col;
        }

    }
}
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class ListColumnSRC : ListColumnBase<ColumnSRC>
    {
        private HashSet<ShapeSheet.SRC> _src_set;

        internal ListColumnSRC() :
            this(0)
        {
        }

        internal ListColumnSRC(int capacity) : base(capacity)
        {
        }

        internal ColumnSRC Add(ShapeSheet.SRC src) => this.Add(src, null);

        internal ColumnSRC Add(ShapeSheet.SRC src, string name)
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
            var col = new ColumnSRC(ordinal, src, name);
            this._items.Add(col);

            this._dic_columns[name] = col;
            this._src_set.Add(src);
            return col;
        }

    }
}
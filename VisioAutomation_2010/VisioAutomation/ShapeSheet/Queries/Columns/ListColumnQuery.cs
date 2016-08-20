using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Columns
{
    public class ListColumnQuery : ListColumnBase<ColumnQuery>
    {
        internal ListColumnQuery() :
            this(0)
        {
        }

        internal ListColumnQuery(int capacity) : base(capacity)
        {
        }

        internal ColumnQuery Add(ShapeSheet.SRC src, string name)
        {
            name = this.fixup_name(name);

            if (this._dic_columns.ContainsKey(name))
            {
                throw new AutomationException("Duplicate Column Name");
            }

            int ordinal = this._columns.Count;
            var col = new ColumnQuery(ordinal, src, name);
            this._columns.Add(col);

            this._dic_columns[name] = col;
            return col;
        }
    }
}
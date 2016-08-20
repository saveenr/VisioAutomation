using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Columns
{
    public class ListColumnSubQuery : ListColumnBase<ColumnSubQuery>
    {
        internal ListColumnSubQuery() :
            base(0)
        {
        }

        internal ListColumnSubQuery(int capacity) : base(capacity)
        {
        }

        public ColumnSubQuery Add(short cell, string name)
        {
            name = this.fixup_name(name);

            if (this._dic_columns.ContainsKey(name))
            {
                throw new AutomationException("Duplicate Column Name");
            }

            int ordinal = this._columns.Count;
            var col = new ColumnSubQuery(ordinal, cell, name);
            this._columns.Add(col);
            return col;
        }
    }
}
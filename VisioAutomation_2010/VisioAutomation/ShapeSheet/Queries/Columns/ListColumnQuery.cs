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

        internal ColumnQuery Add(ShapeSheet.SRC src, string name)
        {
            check_deplicate_src(src);
            string norm_name = this.normalize_name(name);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new ColumnQuery(ordinal, src, norm_name);
            this._items.Add(col);

            this._dic_columns[norm_name] = col;
            this._src_set.Add(src);
            return col;
        }

        private void check_deplicate_src(SRC src)
        {
            if (this._src_set == null)
            {
                this._src_set = new HashSet<ShapeSheet.SRC>();
            }

            if (this._src_set.Contains(src))
            {
                string msg = string.Format("Duplicate SRC({0},{1},{2})", src.Section, src.Row, src.Cell);
                throw new System.ArgumentException(msg);
            }
        }
    }
}
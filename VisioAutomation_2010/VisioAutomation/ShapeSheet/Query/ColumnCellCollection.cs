using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ColumnCellCollection : ColumnCollectionBase<ColumnCell>
    {
        private HashSet<ShapeSheet.SRC> items;

        internal ColumnCellCollection() :
            this(0)
        {
        }

        internal ColumnCellCollection(int capacity) : base(capacity)
        {
        }

        internal ColumnCell Add(ShapeSheet.SRC src, string name)
        {
            check_deplicate_src(src);
            string norm_name = this.normalize_name(name);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new ColumnCell(ordinal, src, norm_name);
            this._items.Add(col);

            this.map_name_to_item[norm_name] = col;
            this.items.Add(src);
            return col;
        }

        private void check_deplicate_src(SRC src)
        {
            if (this.items == null)
            {
                this.items = new HashSet<ShapeSheet.SRC>();
            }

            if (this.items.Contains(src))
            {
                string msg = string.Format("Duplicate SRC({0},{1},{2})", src.Section, src.Row, src.Cell);
                throw new System.ArgumentException(msg);
            }
        }
    }
}
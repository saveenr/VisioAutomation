using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumnList : ColumnListBase<ColumnBase>
    {
        internal CellColumnList() :
            this(0)
        {
        }

        internal CellColumnList(int capacity) : base(capacity)
        {
        }

        public ColumnBase this[VisioAutomation.ShapeSheet.Src src] => this.dic_src_to_col[src];

        public ColumnBase Add(ShapeSheet.Src src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            check_deplicate_src(src);
            string norm_name = this.normalize_name(name);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new ColumnBase(ordinal, norm_name, src);
            this._items.Add(col);

            this.map_name_to_item[norm_name] = col;
            this.dic_src_to_col.Add(src,col);
            return col;
        }


    }
}
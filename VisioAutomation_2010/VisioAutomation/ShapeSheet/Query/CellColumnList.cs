using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumnList : ColumnListBase<CellColumn>
    {
        private Dictionary<ShapeSheet.Src,CellColumn> dic_src_to_col;

        internal CellColumnList() :
            this(0)
        {
        }

        internal CellColumnList(int capacity) : base(capacity)
        {
        }

        public CellColumn this[VisioAutomation.ShapeSheet.Src src] => this.dic_src_to_col[src];

        public CellColumn Add(ShapeSheet.Src src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            check_deplicate_src(src);
            string norm_name = this.normalize_name(name);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new CellColumn(ordinal, src, norm_name);
            this._items.Add(col);

            this.map_name_to_item[norm_name] = col;
            this.dic_src_to_col.Add(src,col);
            return col;
        }

        private void check_deplicate_src(Src src)
        {
            if (this.dic_src_to_col == null)
            {
                this.dic_src_to_col = new Dictionary<ShapeSheet.Src,CellColumn>();
            }

            if (this.dic_src_to_col.ContainsKey(src))
            {
                string msg = string.Format("Duplicate {0}({1},{2},{3})", nameof(Src),src.Section, src.Row, src.Cell);
                throw new System.ArgumentException(msg);
            }
        }
    }
}
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryColumnList : ColumnListBase<SectionQueryColumn>
    {
        internal SectionQueryColumnList() :
            base(0)
        {
        }

        internal SectionQueryColumnList(int capacity) : base(capacity)
        {
        }

        public SectionQueryColumn Add(Src src, string sname)
        {
            check_deplicate_src(src);
                
            string norm_name = this.normalize_name(sname);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new SectionQueryColumn(ordinal, norm_name, src);
            this._items.Add(col);
            this.map_name_to_item[norm_name] = col;

            return col;
        }
    }
}
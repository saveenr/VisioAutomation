using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryColumnList : ColumnListBase<SectionQueryColumn>
    {
        private HashSet<short> hs_cellindex;

        internal SectionQueryColumnList() :
            base(0)
        {
        }

        internal SectionQueryColumnList(int capacity) : base(capacity)
        {
        }

        public SectionQueryColumn Add(Src src, string sname)
        {
            check_duplicate_cellindex(src.Cell);
            string norm_name = this.normalize_name(sname);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new SectionQueryColumn(ordinal, norm_name, src);
            this._items.Add(col);
            this.hs_cellindex.Add(src.Cell);
            this.map_name_to_item[norm_name] = col;

            return col;
        }


        private void check_duplicate_cellindex(short cellindex)
        {
            if (this.hs_cellindex == null)
            {
                this.hs_cellindex = new HashSet<short>();
            }

            if (this.hs_cellindex.Contains(cellindex))
            {
                string msg = string.Format("Duplicate Cell Index: {0}", cellindex);
                throw new System.ArgumentException(msg);
            }
        }
    }
}
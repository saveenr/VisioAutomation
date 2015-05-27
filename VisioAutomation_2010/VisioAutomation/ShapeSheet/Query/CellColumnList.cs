using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumnList : IEnumerable<CellColumn>
    {
        private IList<CellColumn> items { get; }
        private readonly Dictionary<string, CellColumn> dic_columns;
        private HashSet<SRC> hs_src;
        private HashSet<short> hs_cellindex;
        private CellColumnType coltype;

        internal CellColumnList() :
            this(0)
        {
        }

        internal CellColumnList(int capacity)
        {
            this.items = new List<CellColumn>(capacity);
            this.dic_columns = new Dictionary<string, CellColumn>(capacity);
            this.coltype = CellColumnType.Unknown;
        }

        public IEnumerator<CellColumn> GetEnumerator()
        {
            return (this.items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public CellColumn this[int index] => this.items[index];

        public CellColumn this[string name] => this.dic_columns[name];

        public bool Contains(string name) => this.dic_columns.ContainsKey(name);

        internal CellColumn Add(SRC src) => this.Add(src, null);

        internal CellColumn Add(SRC src, string name)
        {
            if (this.coltype == CellColumnType.CellIndex)
            {
                throw new AutomationException("Can't add an SRC if Columns contains CellIndexes");
            }
            this.coltype = CellColumnType.SRC;

            name = this.fixup_name(name);

            if (this.dic_columns.ContainsKey(name))
            {
                throw new AutomationException("Duplicate Column Name");
            }

            if (this.hs_src == null)
            {
                this.hs_src = new HashSet<SRC>();
            }

            if (this.hs_src.Contains(src))
            {
                string msg = "Duplicate SRC";
                throw new AutomationException(msg);
            }

            int ordinal = this.items.Count;
            var col = new CellColumn(ordinal, src, name);
            this.items.Add(col);

            this.dic_columns[name] = col;
            this.hs_src.Add(src);
            return col;
        }

        public CellColumn Add(short cell)
        {
            return this.Add(cell, null);
        }

        public CellColumn Add(short cell, string name)
        {
            if (this.coltype == CellColumnType.SRC)
            {
                throw new AutomationException("Can't add a CellIndex if Columns contains SRCs");
            }

            this.coltype = CellColumnType.CellIndex;

            if (this.hs_cellindex == null)
            {
                this.hs_cellindex = new HashSet<short>();
            }

            if (this.hs_cellindex.Contains(cell))
            {
                string msg = "Duplicate Cell Index";
                throw new AutomationException(msg);
            }

            name = this.fixup_name(name);
            int ordinal = this.items.Count;
            var col = new CellColumn(ordinal, cell, name);
            this.items.Add(col);
            this.hs_cellindex.Add(cell);
            return col;
        }

        private string fixup_name(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = $"Col{this.items.Count}";
            }
            return name;
        }

        public int Count => this.items.Count;
    }
}
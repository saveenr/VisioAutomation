using System;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumnList : IEnumerable<CellColumn>
    {
        private readonly IList<CellColumn> _items;
        private readonly Dictionary<string, CellColumn> _dic_columns;
        private HashSet<SRC> _src_set;
        private HashSet<short> _cellindex_set;
        private CellColumnType _coltype;

        internal CellColumnList() :
            this(0)
        {
        }

        internal CellColumnList(int capacity)
        {
            this._items = new List<CellColumn>(capacity);
            this._dic_columns = new Dictionary<string, CellColumn>(capacity);
            this._coltype = CellColumnType.Unknown;
        }

        public IEnumerator<CellColumn> GetEnumerator()
        {
            return (this._items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public CellColumn this[int index] => this._items[index];

        public CellColumn this[string name] => this._dic_columns[name];

        public bool Contains(string name) => this._dic_columns.ContainsKey(name);

        internal CellColumn Add(SRC src) => this.Add(src, null);

        internal CellColumn Add(SRC src, string name)
        {
            if (this._coltype == CellColumnType.CellIndex)
            {
                throw new AutomationException("Can't add an SRC if Columns contains CellIndexes");
            }
            this._coltype = CellColumnType.SRC;

            name = this.fixup_name(name);

            if (this._dic_columns.ContainsKey(name))
            {
                throw new AutomationException("Duplicate Column Name");
            }

            if (this._src_set == null)
            {
                this._src_set = new HashSet<SRC>();
            }

            if (this._src_set.Contains(src))
            {
                string msg = "Duplicate SRC";
                throw new AutomationException(msg);
            }

            int ordinal = this._items.Count;
            var col = new CellColumn(ordinal, src, name);
            this._items.Add(col);

            this._dic_columns[name] = col;
            this._src_set.Add(src);
            return col;
        }

        public CellColumn Add(short cell)
        {
            return this.Add(cell, null);
        }

        public CellColumn Add(short cell, string name)
        {
            if (this._coltype == CellColumnType.SRC)
            {
                throw new AutomationException("Can't add a CellIndex if Columns contains SRCs");
            }

            this._coltype = CellColumnType.CellIndex;

            if (this._cellindex_set == null)
            {
                this._cellindex_set = new HashSet<short>();
            }

            if (this._cellindex_set.Contains(cell))
            {
                string msg = "Duplicate Cell Index";
                throw new AutomationException(msg);
            }

            name = this.fixup_name(name);
            int ordinal = this._items.Count;
            var col = new CellColumn(ordinal, cell, name);
            this._items.Add(col);
            this._cellindex_set.Add(cell);
            return col;
        }

        private string fixup_name(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = String.Format("Col{0}", this._items.Count);
            }
            return name;
        }

        public int Count => this._items.Count;
    }
}
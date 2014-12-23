using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public class ColumnList : IEnumerable<Column>
        {
            private enum ColumnType
            {
                Unknown,
                SRC,
                CellIndex
            }

            private IList<Column> items { get; set; }
            private Dictionary<string, Column> dic_columns;
            private HashSet<VA.ShapeSheet.SRC> hs_src;
            private HashSet<short> hs_cellindex;
            private ColumnType coltype;

            internal ColumnList() :
                this(0)
            {
            }

            internal ColumnList(int capacity)
            {
                this.items = new List<Column>(capacity);
                this.dic_columns = new Dictionary<string, Column>(capacity);
                this.coltype = ColumnType.Unknown;
            }

            public IEnumerator<Column> GetEnumerator()
            {
                return (this.items).GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public Column this[int index]
            {
                get { return this.items[index]; }
            }

            public Column this[VA.ShapeSheet.Query.CellQuery.Column index]
            {
                get { return this.items[index.Ordinal]; }
            }

            public Column this[string name]
            {
                get { return this.dic_columns[name]; }
            }

            public bool Contains(string name)
            {
                return this.dic_columns.ContainsKey(name);
            }

            public Column Add(SRC src)
            {
                return this.Add(src, null);
            }

            public Column Add(SRC src, string name)
            {
                if (this.coltype == ColumnType.CellIndex)
                {
                    throw new VA.AutomationException("Can't add an SRC if Columns contains CellIndexes");
                }
                this.coltype = ColumnType.SRC;

                name = fixup_name(name);

                if (this.dic_columns.ContainsKey(name))
                {
                    throw new VA.AutomationException("Duplicate Column Name");
                }

                if (this.hs_src == null)
                {
                    this.hs_src = new HashSet<SRC>();
                }

                if (this.hs_src.Contains(src))
                {
                    string msg = string.Format("Duplicate SRC");
                    throw new VA.AutomationException(msg);
                }

                int ordinal = this.items.Count;
                var col = new Column(ordinal, src, name);
                this.items.Add(col);

                this.dic_columns[name] = col;
                this.hs_src.Add(src);
                return col;
            }

            public Column Add(short cell)
            {
                return this.Add(cell, null);
            }

            public Column Add(short cell, string name)
            {
                if (this.coltype == ColumnType.SRC)
                {
                    throw new VA.AutomationException("Can't add a CellIndex if Columns contains SRCs");
                }

                this.coltype = ColumnType.CellIndex;

                if (this.hs_cellindex == null)
                {
                    this.hs_cellindex = new HashSet<short>();
                }

                if (this.hs_cellindex.Contains(cell))
                {
                    string msg = string.Format("Duplicate Cell Index");
                    throw new VA.AutomationException(msg);
                }

                name = fixup_name(name);
                int ordinal = this.items.Count;
                var col = new Column(ordinal, cell, name);
                this.items.Add(col);
                this.hs_cellindex.Add(cell);
                return col;
            }

            private string fixup_name(string name)
            {
                if (string.IsNullOrEmpty(name))
                {
                    name = string.Format("Col{0}", this.items.Count);
                }
                return name;
            }

            public int Count
            {
                get { return this.items.Count; }
            }
        }
    }
}
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Data
{
    public class DataColumnCollection : IEnumerable<DataColumn>
    {
        protected IList<DataColumn> _items;
        protected Dictionary<string, DataColumn> _map_name_to_item;
        protected Dictionary<Core.Src, DataColumn> _dic_src_to_col;
        public IVisio.VisSectionIndices SectionIndex { get; }

        internal DataColumnCollection(IVisio.VisSectionIndices section)
        {
            this._items = new List<DataColumn>();
            this._map_name_to_item = new Dictionary<string, DataColumn>();
            this.SectionIndex = section;
        }


        public IEnumerator<DataColumn> GetEnumerator()
        {
            return (this._items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public DataColumn this[int index] => this._items[index];

        public DataColumn this[string name] => this._map_name_to_item[name];

        public bool Contains(string name) => this._map_name_to_item.ContainsKey(name);

        protected string normalize_name(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                name = string.Format("Col{0}", this._items.Count);
            }

            return name;
        }

        public int Count => this._items.Count;

        protected void check_duplicate_column_name(string name)
        {
            if (this._map_name_to_item.ContainsKey(name))
            {
                throw new System.ArgumentException("Duplicate Column Name");
            }
        }

        protected void check_deplicate_src(Core.Src src)
        {
            if (this._dic_src_to_col == null)
            {
                this._dic_src_to_col = new Dictionary<Core.Src, DataColumn>();
            }

            if (this._dic_src_to_col.ContainsKey(src))
            {
                string msg = string.Format("Duplicate {0}({1},{2},{3})", nameof(Core.Src), src.Section, src.Row,
                    src.Cell);
                throw new System.ArgumentException(msg);
            }
        }

        public DataColumn Add(Core.Src src)
        {
            string name = string.Format("Column{0}", this.Count);
            var col = this.Add(src, name);
            return col;
        }

        public DataColumn Add(Core.Src src, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            check_deplicate_src(src);
            string norm_name = this.normalize_name(name);
            check_duplicate_column_name(norm_name);

            int ordinal = this._items.Count;
            var col = new DataColumn(ordinal, norm_name, src);
            this._items.Add(col);

            this._map_name_to_item[norm_name] = col;
            this._dic_src_to_col.Add(src, col);
            return col;
        }
    }
}
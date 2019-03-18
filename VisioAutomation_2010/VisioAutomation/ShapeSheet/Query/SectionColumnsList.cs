using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumnsList : IEnumerable<SectionColumns>
    {
        private IList<SectionColumns> _list { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SectionColumns> _map_secindex_to_sec_cols; 

        internal SectionColumnsList(int capacity)
        {
            this._list = new List<SectionColumns>(capacity);
            this._map_secindex_to_sec_cols = new Dictionary<IVisio.VisSectionIndices, SectionColumns>(capacity);
        }

        public IEnumerator<SectionColumns> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionColumns this[int index] => this._list[index];

        public SectionColumns Add(IVisio.VisSectionIndices sec_index)
        {
            if (this._map_secindex_to_sec_cols.ContainsKey(sec_index))
            {
                string msg = string.Format("Already contains section index {0} (value={1})",sec_index, (int)sec_index);
                throw new System.ArgumentException(msg);
            }

            var sec_cols = new SectionColumns(sec_index);
            this._list.Add(sec_cols);
            this._map_secindex_to_sec_cols[sec_index] = sec_cols;
            return sec_cols;
        }

        public SectionColumns Add(Src src)
        {
            return this.Add((IVisio.VisSectionIndices)src.Section);
        }

        public int Count => this._list.Count;
    }
}
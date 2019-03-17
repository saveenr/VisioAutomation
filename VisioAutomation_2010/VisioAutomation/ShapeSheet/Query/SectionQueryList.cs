using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryList : IEnumerable<SectionQuery>
    {
        private IList<SectionQuery> _list { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SectionQuery> _map_secindex_to_sectionquery; 

        internal SectionQueryList(int capacity)
        {
            this._list = new List<SectionQuery>(capacity);
            this._map_secindex_to_sectionquery = new Dictionary<IVisio.VisSectionIndices, SectionQuery>(capacity);
        }

        public IEnumerator<SectionQuery> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionQuery this[int index] => this._list[index];

        public SectionQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._map_secindex_to_sectionquery.ContainsKey(section))
            {
                string msg = string.Format("Already contains section index {0} (value={1})",section, (int)section);
                throw new System.ArgumentException(msg);
            }

            var section_query = new SectionQuery(section);
            this._list.Add(section_query);
            this._map_secindex_to_sectionquery[section] = section_query;
            return section_query;
        }

        public SectionQuery Add(Src src)
        {
            return this.Add((IVisio.VisSectionIndices)src.Section);
        }

        public int Count => this._list.Count;
    }
}
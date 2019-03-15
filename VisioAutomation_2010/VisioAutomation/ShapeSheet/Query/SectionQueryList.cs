using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryList : IEnumerable<SectionQuery>
    {
        private IList<SectionQuery> _subqueries { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SectionQuery> _section_set; 

        internal SectionQueryList(int capacity)
        {
            this._subqueries = new List<SectionQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SectionQuery>(capacity);
        }

        public IEnumerator<SectionQuery> GetEnumerator()
        {
            return this._subqueries.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionQuery this[int index] => this._subqueries[index];

        public SectionQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section Index";
                throw new System.ArgumentException(msg);
            }

            int ordinal = this._subqueries.Count;
            var section_query = new SectionQuery(ordinal, section);
            this._subqueries.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public SectionQuery Add(Src src)
        {
            return this.Add((IVisio.VisSectionIndices)src.Section);
        }

        public int Count => this._subqueries.Count;
    }
}
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionSubQueryList : IEnumerable<SectionSubQuery>
    {
        private IList<SectionSubQuery> _subqueries { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SectionSubQuery> _section_set; 

        internal SectionSubQueryList(int capacity)
        {
            this._subqueries = new List<SectionSubQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SectionSubQuery>(capacity);
        }

        public IEnumerator<SectionSubQuery> GetEnumerator()
        {
            return this._subqueries.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionSubQuery this[int index] => this._subqueries[index];

        internal SectionSubQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section Index";
                throw new System.ArgumentException(msg);
            }

            int ordinal = this._subqueries.Count;
            var section_query = new SectionSubQuery(ordinal, section);
            this._subqueries.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this._subqueries.Count;
    }
}
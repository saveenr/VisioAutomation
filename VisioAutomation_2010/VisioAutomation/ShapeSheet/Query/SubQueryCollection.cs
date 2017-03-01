using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryCollection : IEnumerable<SubQuery>
    {
        private IList<SubQuery> _subqueries { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SubQuery> _section_set; 

        internal SubQueryCollection(int capacity)
        {
            this._subqueries = new List<SubQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SubQuery>(capacity);
        }

        public IEnumerator<SubQuery> GetEnumerator()
        {
            return this._subqueries.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SubQuery this[int index] => this._subqueries[index];

        internal SubQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section Index";
                throw new System.ArgumentException(msg);
            }

            int ordinal = this._subqueries.Count;
            var section_query = new SubQuery(ordinal, section);
            this._subqueries.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this._subqueries.Count;
    }
}
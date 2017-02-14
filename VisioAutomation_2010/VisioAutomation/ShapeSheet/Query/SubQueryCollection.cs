using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SubQueryCollection : IEnumerable<SubQuery>
    {
        private IList<SubQuery> items { get; }

        private readonly Dictionary<IVisio.VisSectionIndices,SubQuery> _section_set; 

        internal SubQueryCollection(int capacity)
        {
            this.items = new List<SubQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SubQuery>(capacity);
        }

        public IEnumerator<SubQuery> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SubQuery this[int index] => this.items[index];

        internal SubQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section";
                throw new System.ArgumentException(msg);
            }

            int ordinal = this.items.Count;
            var section_query = new SubQuery(ordinal, section);
            this.items.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this.items.Count;
    }
}
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries
{
    public class ListSubQuery : IEnumerable<SubQuery>
    {
        private IList<SubQuery> Items { get; }
        private readonly Dictionary<IVisio.VisSectionIndices,SubQuery> _section_set; 

        internal ListSubQuery(int capacity)
        {
            this.Items = new List<SubQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SubQuery>(capacity);
        }

        public IEnumerator<SubQuery> GetEnumerator()
        {
            return this.Items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SubQuery this[int index] => this.Items[index];

        internal SubQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section";
                throw new AutomationException(msg);
            }

            int ordinal = this.Items.Count;
            var section_query = new SubQuery(ordinal, section);
            this.Items.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this.Items.Count;
    }
}
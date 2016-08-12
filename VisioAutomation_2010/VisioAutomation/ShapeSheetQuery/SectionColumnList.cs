using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheetQuery
{
    public class SectionColumnList : IEnumerable<SectionSubQuery>
    {
        private IList<SectionSubQuery> Items { get; }
        private readonly Dictionary<IVisio.VisSectionIndices,SectionSubQuery> _section_set; 

        internal SectionColumnList(int capacity)
        {
            this.Items = new List<SectionSubQuery>(capacity);
            this._section_set = new Dictionary<IVisio.VisSectionIndices, SectionSubQuery>(capacity);
        }

        public IEnumerator<SectionSubQuery> GetEnumerator()
        {
            return this.Items.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionSubQuery this[int index] => this.Items[index];

        internal SectionSubQuery Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section";
                throw new AutomationException(msg);
            }

            int ordinal = this.Items.Count;
            var section_query = new SectionSubQuery(ordinal, section);
            this.Items.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this.Items.Count;
    }
}
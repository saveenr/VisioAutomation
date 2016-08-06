using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumnList : System.Collections.Generic.IEnumerable<SectionColumn>
    {
        private System.Collections.Generic.IList<SectionColumn> Items { get; }
        private readonly System.Collections.Generic.Dictionary<IVisio.VisSectionIndices,SectionColumn> _section_set; 

        internal SectionColumnList(int capacity)
        {
            this.Items = new System.Collections.Generic.List<SectionColumn>(capacity);
            this._section_set = new System.Collections.Generic.Dictionary<IVisio.VisSectionIndices, SectionColumn>(capacity);
        }

        public System.Collections.Generic.IEnumerator<SectionColumn> GetEnumerator()
        {
            return (this.Items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionColumn this[int index] => this.Items[index];

        internal SectionColumn Add(IVisio.VisSectionIndices section)
        {
            if (this._section_set.ContainsKey(section))
            {
                string msg = "Duplicate Section";
                throw new AutomationException(msg);
            }

            int ordinal = this.Items.Count;
            var section_query = new SectionColumn(ordinal, section);
            this.Items.Add(section_query);
            this._section_set[section] = section_query;
            return section_query;
        }

        public int Count => this.Items.Count;
    }
}
namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumnList : System.Collections.Generic.IEnumerable<SectionColumn>
    {
        private System.Collections.Generic.IList<SectionColumn> items { get; }
        private readonly System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Visio.VisSectionIndices,SectionColumn> hs_section; 

        internal SectionColumnList(int capacity)
        {
            this.items = new System.Collections.Generic.List<SectionColumn>(capacity);
            this.hs_section = new System.Collections.Generic.Dictionary<Microsoft.Office.Interop.Visio.VisSectionIndices, SectionColumn>(capacity);
        }

        public System.Collections.Generic.IEnumerator<SectionColumn> GetEnumerator()
        {
            return (this.items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionColumn this[int index] => this.items[index];

        internal SectionColumn Add(Microsoft.Office.Interop.Visio.VisSectionIndices section)
        {
            if (this.hs_section.ContainsKey(section))
            {
                string msg = "Duplicate Section";
                throw new AutomationException(msg);
            }

            int ordinal = this.items.Count;
            var section_query = new SectionColumn(ordinal, section);
            this.items.Add(section_query);
            this.hs_section[section] = section_query;
            return section_query;
        }

        public int Count => this.items.Count;
    }
}
using System;
using VA = VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumnList : IEnumerable<SectionColumn>
    {
        private IList<SectionColumn> items { get; set; }
        private readonly Dictionary<IVisio.VisSectionIndices,SectionColumn> hs_section; 

        internal SectionColumnList(int capacity)
        {
            this.items = new List<SectionColumn>(capacity);
            this.hs_section = new Dictionary<IVisio.VisSectionIndices, SectionColumn>(capacity);
        }

        public IEnumerator<SectionColumn> GetEnumerator()
        {
            return (this.items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public SectionColumn this[int index]
        {
            get { return this.items[index]; }
        }

        internal SectionColumn Add(IVisio.VisSectionIndices section)
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

        public int Count
        {
            get { return this.items.Count; }
        }
    }
}
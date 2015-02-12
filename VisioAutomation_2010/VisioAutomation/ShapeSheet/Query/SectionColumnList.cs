using System;
using VA = VisioAutomation;
using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionColumnList : IEnumerable<CellQuery.SectionColumn>
    {
        private IList<CellQuery.SectionColumn> items { get; set; }
        private readonly Dictionary<IVisio.VisSectionIndices,CellQuery.SectionColumn> hs_section; 
 
        internal SectionColumnList() :
            this(0)
        {
        }

        internal SectionColumnList(int capacity)
        {
            this.items = new List<CellQuery.SectionColumn>(capacity);
            this.hs_section = new Dictionary<IVisio.VisSectionIndices, CellQuery.SectionColumn>(capacity);
        }

        public IEnumerator<CellQuery.SectionColumn> GetEnumerator()
        {
            return (this.items).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public CellQuery.SectionColumn this[int index]
        {
            get { return this.items[index]; }
        }

        internal CellQuery.SectionColumn Add(IVisio.VisSectionIndices section)
        {
            if (this.hs_section.ContainsKey(section))
            {
                string msg = String.Format("Duplicate Section");
                throw new AutomationException(msg);
            }

            int ordinal = items.Count;
            var section_query = new CellQuery.SectionColumn(ordinal, section);
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
using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionQueryList : IEnumerable<SectionQuery>
       {
           private IList<SectionQuery> items { get; set; }
           private readonly CellQuery parent;
           private readonly Dictionary<IVisio.VisSectionIndices,SectionQuery> hs_section; 
 
           internal SectionQueryList(CellQuery parent) :
               this(parent,0)
           {
           }

           internal SectionQueryList(CellQuery parent,int capacity)
           {
               this.items = new List<SectionQuery>(capacity);
               this.parent = parent;
               this.hs_section = new Dictionary<IVisio.VisSectionIndices, SectionQuery>(capacity);
           }

           public IEnumerator<SectionQuery> GetEnumerator()
           {
               return (this.items).GetEnumerator();
           }

           System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
           {
               return GetEnumerator();
           }

           public SectionQuery this[int index]
           {
               get { return this.items[index]; }
           }

           public SectionQuery Add(IVisio.VisSectionIndices section)
           {
               if (this.hs_section.ContainsKey(section))
               {
                   string msg = string.Format("Duplicate Section");
                   throw new AutomationException(msg);
               }

               int ordinal = items.Count;
               var section_query = new SectionQuery(this.parent, ordinal, section);
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

}
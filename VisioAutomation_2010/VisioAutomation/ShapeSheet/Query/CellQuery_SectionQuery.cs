using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {

       public class SectionQuery
       {
           public CellQuery Parent { get; private set; }
           public short SectionIndex { get; private set; }
           public List<Column> Columns { get; private set; }
           public int Ordinal { get; private set; }

           public SectionQuery(CellQuery parent, int ordinal, short section)
           {
               this.Parent = parent;
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.Columns = new List<Column>();
           }

           public static implicit operator int(SectionQuery m)
           {
               return m.Ordinal;
           }

           public Column AddColumn(SRC src, string name)
           {
               this.Parent.CheckNotFrozen();

               if (string.IsNullOrEmpty(name))
               {
                   name = string.Format("Col{0}", this.Columns.Count);
               }
               
               int ordinal = this.Columns.Count;
               if (src.Section != this.SectionIndex)
               {
                   throw new VA.AutomationException("SRC's Section does not match");
               }
               var col = new Column(ordinal, src, name);
               this.Columns.Add(col);
               return col;
           }

           public Column AddColumn(short cell, string name)
           {
               this.Parent.CheckNotFrozen();

               if (string.IsNullOrEmpty(name))
               {
                   name = string.Format("Col{0}", this.Columns.Count);
               }

               int ordinal = this.Columns.Count;
               var col = new Column(ordinal, cell, name);
               this.Columns.Add(col);
               return col;
           }
       }
    }
}
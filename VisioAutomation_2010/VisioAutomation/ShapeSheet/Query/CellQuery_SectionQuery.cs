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
           public ColumnList Columns { get; private set; }
           public int Ordinal { get; private set; }

           private HashSet<short> hs_cellidnex; 

           internal SectionQuery(CellQuery parent, int ordinal, short section)
           {
               this.Parent = parent;
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.Columns = new ColumnList();
               this.hs_cellidnex = new HashSet<short>();
           }

           public Column AddColumn(SRC src, string name)
           {
               this.Parent.CheckNotFrozen();
               return this.Columns.Add(src, name);
           }

           public Column AddColumn(short cell, string name)
           {
               this.Parent.CheckNotFrozen();

               if (this.hs_cellidnex.Contains(cell))
               {
                   throw new VA.AutomationException("Duplicate Cell");
               }

               return this.Columns.Add(cell, name);
           }
       }
    }
}
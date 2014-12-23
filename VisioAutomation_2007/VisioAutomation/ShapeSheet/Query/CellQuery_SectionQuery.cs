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
           public IVisio.VisSectionIndices SectionIndex { get; private set; }
           public ColumnList Columns { get; private set; }
           public int Ordinal { get; private set; }

           internal SectionQuery(CellQuery parent, int ordinal, IVisio.VisSectionIndices section)
           {
               this.Parent = parent;
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.Columns = new ColumnList();
           }
       }
    }
}
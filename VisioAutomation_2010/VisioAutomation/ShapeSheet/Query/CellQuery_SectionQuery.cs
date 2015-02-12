using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionQuery
       {
           public IVisio.VisSectionIndices SectionIndex { get; private set; }
           public CellColumnList CellColumns { get; private set; }
           public int Ordinal { get; private set; }

           internal SectionQuery(int ordinal, IVisio.VisSectionIndices section)
           {
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.CellColumns = new CellColumnList();
           }

           public Column AddCell(SRC src, string name)
           {
               var col = this.CellColumns.Add(src, name);
               return col;
           }

           public Column AddCell(short cell)
           {
               return this.AddCell(cell, null);
           }

           public Column AddCell(short cell, string name)
           {
               var col = this.CellColumns.Add(cell, name);
               return col;
           }
       }
    }
}
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionColumn
       {
           public IVisio.VisSectionIndices SectionIndex { get; private set; }
           public CellColumnList CellColumns { get; private set; }
           public int Ordinal { get; private set; }

           internal SectionColumn(int ordinal, IVisio.VisSectionIndices section)
           {
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.CellColumns = new CellColumnList();
           }

           public CellColumn AddCell(SRC src, string name)
           {
               var col = this.CellColumns.Add(src, name);
               return col;
           }

           public CellColumn AddCell(short cell, string name)
           {
               var col = this.CellColumns.Add(cell, name);
               return col;
           }
       }
    }
}
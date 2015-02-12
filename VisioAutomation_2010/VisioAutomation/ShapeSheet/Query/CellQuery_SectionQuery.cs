using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionQuery
       {
           public IVisio.VisSectionIndices SectionIndex { get; private set; }
           public ColumnList Columns { get; private set; }
           public int Ordinal { get; private set; }

           internal SectionQuery(int ordinal, IVisio.VisSectionIndices section)
           {
               this.Ordinal = ordinal;
               this.SectionIndex = section;
               this.Columns = new ColumnList();
           }
       }
    }
}
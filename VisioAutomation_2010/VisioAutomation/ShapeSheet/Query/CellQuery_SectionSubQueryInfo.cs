using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionSubQueryInfo
       {
           public SectionSubQuery SectionSubQuery { get; private set; }
           public short ShapeID { get; private set; }
           public List<short> RowIndexes { get; private set; }

           public SectionSubQueryInfo(SectionSubQuery sq, short shapeid, int numrows)
           {
               this.SectionSubQuery = sq;
               this.ShapeID = shapeid;
               this.RowIndexes = Enumerable.Range(0, numrows).Select(i => (short)i).ToList();
           }
       }
    }
}
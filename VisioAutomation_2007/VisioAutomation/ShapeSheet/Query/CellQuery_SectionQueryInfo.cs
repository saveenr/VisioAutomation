using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionQueryInfo
       {
           public SectionQuery SectionQuery { get; private set; }
           public short ShapeID { get; private set; }
           public int RowCount  { get; private set; }

           internal SectionQueryInfo(SectionQuery sq, short shapeid, int numrows)
           {
               this.SectionQuery = sq;
               this.ShapeID = shapeid;
               this.RowCount = numrows;
           }

           public IEnumerable<int> RowIndexes
           {
               get { return Enumerable.Range(0, this.RowCount); }
           }
       }
    }
}
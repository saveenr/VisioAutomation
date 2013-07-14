using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class QueryResult<T>
       {
           public int ShapeID { get; private set; }
           public T[] Cells { get; internal set; }
           public List<SectionResult<T>> SectionCells { get; internal set; }

           public QueryResult(int sid)
           {
               this.ShapeID = sid;
           }
       }
    }
}
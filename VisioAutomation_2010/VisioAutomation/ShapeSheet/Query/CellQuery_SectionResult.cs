using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionResult<T>
       {
           public VA.ShapeSheet.Query.CellQuery.SectionSubQuery Query { get; internal set; }
           public List<T[]> Rows { get; internal set; }
       }
    }
}
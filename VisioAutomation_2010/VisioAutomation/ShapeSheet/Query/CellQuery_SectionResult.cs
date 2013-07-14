using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class SectionResult<T> : IEnumerable<T[]>
       {
           public VA.ShapeSheet.Query.CellQuery.SectionSubQuery Query { get; internal set; }
           private List<T[]> Rows;

           public SectionResult()
           {
               this.Rows = new List<T[]>();
           }

           public IEnumerator<T[]> GetEnumerator()
           {
               return this.Rows.GetEnumerator();
           }

           System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
           {
               return GetEnumerator();
           }


           public T[] this[int index]
           {
               get { return this.Rows[index]; }
           }

           internal void Add(T[] item)
           {
               this.Rows.Add(item);
           }

           public int Count
           {
               get { return this.Rows.Count; }
           }
       }
    }
}
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
           public SectionQuery Query { get; internal set; }
           private List<T[]> items;

           internal SectionResult(int capacity)
           {
               this.items = new List<T[]>(capacity);
           }

           public IEnumerator<T[]> GetEnumerator()
           {
               return this.items.GetEnumerator();
           }

           System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
           {
               return GetEnumerator();
           }

           public T[] this[int index]
           {
               get { return this.items[index]; }
           }

           internal void Add(T[] item)
           {
               this.items.Add(item);
           }

           public int Count
           {
               get { return this.items.Count; }
           }
       }
    }
}
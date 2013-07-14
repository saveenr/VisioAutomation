using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
       public class ColumnList : IEnumerable<Column>
       {
           private IList<Column> items { get; set; }
           
           internal ColumnList() :
               this(0)
           {
           }

           internal ColumnList(int capacity)
           {
               this.items = new List<Column>(capacity);
           }

           public IEnumerator<Column> GetEnumerator()
           {
               return (this.items).GetEnumerator();
           }

           System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
           {
               return GetEnumerator();
           }

           public Column this[int index]
           {
               get { return this.items[index]; }
           }
           public Column this[VA.ShapeSheet.Query.CellQuery.Column index]
           {
               get { return this.items[index.Ordinal]; }
           }

           public Column Add(SRC src, string name)
           {
               name = GetName(name);
               int ordinal = this.items.Count;
               var col = new Column(ordinal, src, name);
               this.items.Add(col);
               return col;
           }

           public Column Add(short cell, string name)
           {
               name = GetName(name);
               int ordinal = this.items.Count;
               var col = new Column(ordinal, cell, name);
               this.items.Add(col);
               return col;
           }

           private string GetName(string name)
           {
               if (string.IsNullOrEmpty(name))
               {
                   name = string.Format("Col{0}", this.items.Count);
               }
               return name;
           }

           public int Count
           {
               get { return this.items.Count; }
           }
       }
    }
}
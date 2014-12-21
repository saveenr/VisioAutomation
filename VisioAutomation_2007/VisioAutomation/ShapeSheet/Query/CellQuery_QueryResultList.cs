using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
   public partial class CellQuery
    {
        public class QueryResultList<T> : IEnumerable<QueryResult<T>>
        {
            List<QueryResult<T>> Items;

            internal QueryResultList()
            {
                this.Items = new List<QueryResult<T>>();
            }

            public IEnumerator<QueryResult<T>> GetEnumerator()
            {
                return this.Items.GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public QueryResult<T> this[int index]
            {
                get { return this.Items[index]; }
            }

            internal void Add(QueryResult<T> item)
            {
                this.Items.Add(item);
            }

            public int Count
            {
                get { return this.Items.Count; }
            }
        }
    }
}
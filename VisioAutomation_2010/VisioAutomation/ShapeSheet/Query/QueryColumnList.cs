using System.Collections.Generic;
using System.Collections;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryColumnList : IEnumerable<QueryColumn>
    {
        private IList<QueryColumn> _columns;

        public QueryColumnList()
        {
            this._columns = new List<QueryColumn>();
        }

        public int Count
        {
            get { return this._columns.Count; }
        }

        public void Add(QueryColumn item)
        {
            this._columns.Add(item);
        }

        public QueryColumn this[int index]
        {
            get { return this._columns[index];  }
        }

        public IEnumerator<QueryColumn> GetEnumerator()
        {
            foreach (var i in this._columns)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
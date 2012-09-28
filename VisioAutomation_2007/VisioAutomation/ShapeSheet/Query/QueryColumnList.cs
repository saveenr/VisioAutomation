using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{

    public class QueryColumnList 
    {
        private IList<QueryColumn> _columns;

        public QueryColumnList()
        {
            this._columns = new List<QueryColumn>();
        }

        public IEnumerable<QueryColumn> Items
        {
            get { return this._columns; }
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
    }
}
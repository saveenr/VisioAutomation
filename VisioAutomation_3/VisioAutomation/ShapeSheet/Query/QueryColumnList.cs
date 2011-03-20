using System;
using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{

    public class QueryColumnList<TCol> where TCol : QueryColumn
    {
        private IList<TCol> _columns;

        public QueryColumnList()
        {
            this._columns = new List<TCol>();
        }

        public IEnumerable<TCol> Items
        {
            get { return this._columns; }
        }

        public int Count
        {
            get { return this._columns.Count; }
        }

        public void Add(TCol item)
        {
            this._columns.Add(item);
        }

        public TCol this[int index]
        {
            get { return this._columns[index];  }
        }
    }
}
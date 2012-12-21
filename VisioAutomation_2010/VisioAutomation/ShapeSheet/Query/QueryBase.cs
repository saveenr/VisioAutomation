using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryBase
    {
        private QueryColumnList _columns;

        internal QueryBase()
        {
            this._columns = new QueryColumnList();
        }

        public QueryColumnList Columns
        {
            get { return this._columns; }
        }

        protected void AddColumn(QueryColumn column)
        {
            if (column == null)
            {
                throw new System.ArgumentNullException("column");
            }

            this._columns.Add(column);
        }

        protected IList<IVisio.VisUnitCodes> CreateUnitCodeArrayForRows(int rowcount)
        {
            if (rowcount<1)
            {
                throw new AutomationException("Must have at least 1 row");
            }

            int n = this.Columns.Count*rowcount;
            var unitcodes = new IVisio.VisUnitCodes[n];
            for (int i = 0; i < n; i++)
            {
                unitcodes[i] = this.Columns[i%this.Columns.Count].UnitCode;
            }

            return unitcodes;
        }
    }
}
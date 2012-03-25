using System;
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
                throw new ArgumentNullException("column");
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
            var a = new IVisio.VisUnitCodes[n];
            for (int i = 0; i < n; i++)
            {
                a[i] = this.Columns[i%this.Columns.Count].UnitCode;
            }

            return a;
        }

        protected void validate_unitcodes(IList<IVisio.VisUnitCodes> unitcodes, int total_cells)
        {
            // ensure that the number of unit codes is equal to total number of cells being retrieved 

            if (unitcodes == null)
            {
                throw new AutomationException("unitcodes must not be null");
            }

            if (unitcodes.Count != total_cells)
            {
                string msg = string.Format("Expected {0} unitcodes", total_cells);
                throw new AutomationException(msg);
            }
        }
    }
}
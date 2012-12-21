using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryBase
    {
        public List<QueryColumn> Columns { get; private set; }

        internal QueryBase()
        {
            this.Columns = new List<QueryColumn>();
        }

        protected void AddColumn(QueryColumn column)
        {
            if (column == null)
            {
                throw new System.ArgumentNullException("column");
            }

            this.Columns.Add(column);
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
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryColumn
    {
        public string Name { get; private set; }
        public IVisio.VisUnitCodes UnitCode { get; set; }

        protected QueryColumn(int ordinal, string name)
        {
            this.Ordinal = ordinal;
            this.UnitCode = IVisio.VisUnitCodes.visNoCast;

            this.Name = name ?? string.Format("Column{0}", ordinal);
        }

        public int Ordinal { get; private set; }
    }
}
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class QueryColumn
    {
        public string Name { get; private set; }
        public IVisio.VisUnitCodes UnitCode { get; set; }
        public SRC SRC { get; protected set; }

        protected QueryColumn(int ordinal, string name)
        {
            this.Ordinal = ordinal;
            this.UnitCode = IVisio.VisUnitCodes.visNoCast;

            this.Name = name ?? string.Format("Column{0}", ordinal);
        }

        public int Ordinal { get; private set; }

        internal QueryColumn(int ordinal, short cell, string name) :
            this(ordinal,name)
        {
            this.SRC = new VA.ShapeSheet.SRC(-1,-1,cell);
        }

        internal QueryColumn(int ordinal, SRC src, string name) :
            this(ordinal,name)
        {
            this.SRC = src;
        }

        public static implicit operator int(QueryColumn m)
        {
            return m.Ordinal;
        }


    }
}
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public class Column
        {
            public string Name { get; private set; }
            public IVisio.VisUnitCodes UnitCode { get; set; }
            public SRC SRC { get; protected set; }

            protected Column(int ordinal, string name)
            {
                if (string.IsNullOrEmpty(name))
                {
                    throw new System.ArgumentException("name");
                }

                this.Ordinal = ordinal;
                this.UnitCode = IVisio.VisUnitCodes.visNoCast;
            }

            public int Ordinal { get; private set; }

            internal Column(int ordinal, short cell, string name) :
                this(ordinal, name)
            {
                this.SRC = new VA.ShapeSheet.SRC(-1, -1, cell);
            }

            internal Column(int ordinal, SRC src, string name) :
                this(ordinal, name)
            {
                this.SRC = src;
            }

            public static implicit operator int(Column m)
            {
                return m.Ordinal;
            }
        }
    }
}
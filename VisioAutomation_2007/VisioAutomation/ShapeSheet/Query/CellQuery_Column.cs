using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public partial class CellQuery
    {
        public class Column
        {
            public string Name { get; private set; }
            public SRC SRC { get; protected set; }
            public IVisio.VisUnitCodes UnitCode { get; set; }
            
            protected Column(int ordinal, string name)
            {
                if (string.IsNullOrEmpty(name))
                {
                    throw new System.ArgumentException("name");
                }
                this.Name = name;
                this.Ordinal = ordinal;
            }

            public int Ordinal { get; private set; }

            internal Column(int ordinal, short cell, string name) :
                this(ordinal, name)
            {
                const short sec = -1;
                const short row = -1;
                this.SRC = new VA.ShapeSheet.SRC(sec, row, cell);
            }

            internal Column(int ordinal, SRC src, string name) :
                this(ordinal, name)
            {
                this.SRC = src;
            }
        }
    }
}
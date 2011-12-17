using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryColumn : QueryColumn
    {
        public SRC SRC { get; protected set; }


        internal CellQueryColumn(int ordinal, SRC src, string name) :
            base(ordinal,name)
        {
            this.SRC = src;
        }
    }
}
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class SectionQueryColumn : QueryColumn
    {
        public short Cell { get; protected set; }

        internal SectionQueryColumn(int ordinal, short cell, string name) :
            base(ordinal,name)
        {
            this.Cell = cell;
        }
    }
}
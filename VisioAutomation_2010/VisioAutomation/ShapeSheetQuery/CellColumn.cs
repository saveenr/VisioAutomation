using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class CellColumn
    {
        public string Name { get; private set; }
        public VisioAutomation.ShapeSheet.SRC SRC { get; protected set; }
        public IVisio.VisUnitCodes UnitCode { get; set; }
        public int Ordinal { get; }
            
        protected CellColumn(int ordinal, string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }
 
            this.Name = name;
            this.Ordinal = ordinal;
        }

        internal CellColumn(int ordinal, short cell, string name) :
            this(ordinal, name)
        {
            const short sec = -1;
            const short row = -1;
            this.SRC = new VisioAutomation.ShapeSheet.SRC(sec, row, cell);
        }

        internal CellColumn(int ordinal, VisioAutomation.ShapeSheet.SRC src, string name) :
            this(ordinal, name)
        {
            this.SRC = src;
        }

        public static implicit operator int (CellColumn col)
        {
            return col.Ordinal;
        }
    }
}
using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellColumn
    {
        public string Name { get; private set; }
        public SRC SRC { get; protected set; }
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
            this.SRC = new SRC(sec, row, cell);
        }

        internal CellColumn(int ordinal, SRC src, string name) :
            this(ordinal, name)
        {
            this.SRC = src;
        }

        static public implicit operator int (CellColumn col)
        {
            return col.Ordinal;
        }
    }
}
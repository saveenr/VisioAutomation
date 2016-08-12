using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery
{
    public class CellColumnBase
    {
        public string Name { get; protected set; }
        public IVisio.VisUnitCodes UnitCode { get; set; }
        public int Ordinal { get; protected set; }

        protected CellColumnBase(int ordinal, string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }

            this.Name = name;
            this.Ordinal = ordinal;
        }

        public static implicit operator int(CellColumnBase col)
        {
            return col.Ordinal;
        }
    }

    public class CellColumn : CellColumnBase
    {
        public ShapeSheet.SRC SRC { get; protected set; }

        internal CellColumn(int ordinal, ShapeSheet.SRC src, string name) :
            base(ordinal, name)
        {
            this.SRC = src;
        }

    }

    public class SubQueryCellColumn : CellColumnBase
    {
        public short CellIndex;

        internal SubQueryCellColumn(int ordinal, short cell, string name) :
            base(ordinal, name)
        {
            this.CellIndex = cell;
        }
    }

}
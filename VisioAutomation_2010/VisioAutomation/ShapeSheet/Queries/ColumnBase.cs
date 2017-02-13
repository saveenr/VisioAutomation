using System;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Queries
{
    public class ColumnBase
    {
        public string Name { get; protected set; }
        public IVisio.VisUnitCodes UnitCode { get; set; }
        public int Ordinal { get; protected set; }

        protected ColumnBase(int ordinal, string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }

            this.Name = name;
            this.Ordinal = ordinal;
            this.UnitCode = IVisio.VisUnitCodes.visNoCast;
        }

        public static implicit operator int(ColumnBase col)
        {
            return col.Ordinal;
        }
    }
}
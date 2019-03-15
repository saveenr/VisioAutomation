    using System;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ColumnBase
    {
        public string Name { get; protected set; }
        public int Ordinal { get; protected set; }
        protected ColumnBase(int ordinal, string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }

            this.Name = name;
            this.Ordinal = ordinal;
        }

        public static implicit operator int(ColumnBase col)
        {
            return col.Ordinal;
        }
    }
}
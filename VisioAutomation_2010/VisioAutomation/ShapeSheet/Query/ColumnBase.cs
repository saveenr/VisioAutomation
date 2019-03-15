    using System;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ColumnBase
    {
        public string Name { get; protected set; }
        public int Ordinal { get; protected set; }

        public readonly ShapeSheet.Src Src;

        protected ColumnBase(int ordinal, string name, ShapeSheet.Src src) 
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }

            this.Src = src;
            this.Name = name;
            this.Ordinal = ordinal;
        }
        protected ColumnBase(int ordinal, string name)
        {
        }

        public static implicit operator int(ColumnBase col)
        {
            return col.Ordinal;
        }
    }
}
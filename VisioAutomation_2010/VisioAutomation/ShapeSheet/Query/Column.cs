using System;

namespace VisioAutomation.ShapeSheet.Query
{
    public class Column
    {
        public string Name { get; protected set; }
        public int Ordinal { get; protected set; }

        public ShapeSheet.Src Src { get; }

        public Column(int ordinal, string name, ShapeSheet.Src src) 
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name");
            }

            this.Src = src;
            this.Name = name;
            this.Ordinal = ordinal;
        }
        protected Column(int ordinal, string name)
        {
        }

        public static implicit operator int(Column col)
        {
            return col.Ordinal;
        }


    }
}
using System;

namespace ExcelUtil
{
    public class ColumnDefinition
    {
        public string Name { get; private set; }
        public Type Type { get; private set; }

        public ColumnDefinition(string name, Type type)
        {
            this.Name = name;
            this.Type = type;
        }
    }
}
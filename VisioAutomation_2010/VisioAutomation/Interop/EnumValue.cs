using System;

namespace VisioAutomation.Interop
{
    public class EnumValue
    {
        public string Name { get; }
        public int Value { get; }

        public EnumValue(string name, int value)
        {
            this.Name = name;
            this.Value = value;
        }

        public override string ToString()
        {
            return string.Format("{0},{1}", this.Name, this.Value);
        }
    }
}
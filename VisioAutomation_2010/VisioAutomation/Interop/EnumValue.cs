namespace VisioAutomation.Interop
{
    public class EnumValue
    {
        public string Name { get; private set; }
        public int Value { get; private set; }

        public EnumValue(string name, int value)
        {
            this.Name = name;
            this.Value = value;
        }

        public override string ToString()
        {
            return $"{this.Name},{this.Value}";
        }
    }
}
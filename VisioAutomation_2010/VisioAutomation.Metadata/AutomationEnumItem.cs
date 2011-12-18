namespace VisioAutomation.Metadata
{
    public class AutomationEnumItem
    {
        public string Name;
        public int Value;

        internal AutomationEnumItem(string name, int value)
        {
            this.Name = name;
            this.Value = value;
        }
    }
}
namespace VisioAutomation.Metadata
{
    public class AutomationConstant
    {
        public string Enum { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }

        public int GetValueAsInt()
        {
            return int.Parse(this.Value);
        }
    }
}
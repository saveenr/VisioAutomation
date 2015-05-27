namespace VisioAutomation.Application.Logging
{
    public class LogRecord
    {
        public string Type;
        public string SubType;
        public string Context;
        public string Description;

        public override string ToString()
        {
            return $"{this.Type}:{this.SubType}";
        }
    }
}
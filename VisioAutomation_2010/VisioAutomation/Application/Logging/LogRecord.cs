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
            return string.Format("{0}:{1}", this.Type, this.SubType);
        }
    }
}
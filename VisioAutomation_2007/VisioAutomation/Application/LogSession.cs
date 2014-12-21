using System.Collections.Generic;

namespace VisioAutomation.Application.Logging
{
    public class LogSession
    {
        public string StartLine;
        public string EndLine;

        public List<LogRecord> Records;

        public LogSession()
        {
            this.Records = new List<LogRecord>();
        }
    }
}
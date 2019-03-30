using System.Collections.Generic;

namespace VisioAutomation.Application.Logging
{
    public class LogSession
    {
        public string StartLine;
        public string EndLine;
        public string FileType;
        public string Source;
        public string StartTimeRaw;
        public System.DateTime StartTime;
        public List<LogRecord> LogRecords;

        internal LogSession()
        {
            this.LogRecords = new List<LogRecord>();
        }
    }
}
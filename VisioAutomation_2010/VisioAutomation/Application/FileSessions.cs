using System.Collections.Generic;

namespace VisioAutomation.Application.Logging
{
    public class FileSessions
    {
        public string StartLine;
        public string EndLine;
        public string FileType;
        public string Source;

        public string StartTimeRaw;
        public System.DateTime StartTime;

        public List<LogRecord> Records;

        public FileSessions()
        {
            this.Records = new List<LogRecord>();
        }
    }
}
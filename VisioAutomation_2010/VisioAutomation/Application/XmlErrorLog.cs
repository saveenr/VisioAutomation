using System.Collections.Generic;

namespace VisioAutomation.Application.Logging
{
    public class XmlErrorLog
    {
        public string Source;
        public List<LogSession> Sessions;

        public XmlErrorLog(string filename)
        {
            this.Sessions = new List<LogSession>();

            if (!System.IO.File.Exists(filename))
            {
                return;
            }

            var state = LogState.Start;

            var lines = GetLinesSharedRead(filename);

            foreach (var rawline in lines)
            {
                string line = rawline.Trim();

                if (state == LogState.Start)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }
                    else if (line.StartsWith("Open"))
                    {
                        // do something
                    }
                    else if (line.StartsWith("Source:"))
                    {
                        this.Source = line.Substring("Source:".Length).Trim();
                    }
                    else if (line.StartsWith("Insert"))
                    {
                        var session = new LogSession();
                        session.StartLine = line;
                        this.Sessions.Add(session);
                        state = LogState.InSession;
                    }
                    else if (line.EndsWith("Begin Session"))
                    {
                        var session = new LogSession();
                        session.StartLine = line;
                        this.Sessions.Add(session);
                        state = LogState.InSession;
                    }
                    else
                    {
                        throw new System.ArgumentException();
                    }
                }
                else if (state == LogState.InSession)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }
                    if (line.EndsWith("End Session"))
                    {
                        var session = this.Sessions[this.Sessions.Count - 1];
                        session.EndLine = line;
                        state = LogState.Start;
                    }
                    else if (line.StartsWith("Open"))
                    {
                        var session = this.Sessions[this.Sessions.Count - 1];
                        session.EndLine = line;
                        state = LogState.Start;
                    }
                    else if (line.StartsWith("Insert"))
                    {
                        var session = this.Sessions[this.Sessions.Count - 1];
                        session.EndLine = line;
                        state = LogState.Start;
                    }
                    else if (line.StartsWith("["))
                    {
                        var rec = new LogRecord();
                        int n = line.IndexOf(']');
                        if (n < 2)
                        {
                            throw new System.ArgumentException();
                        }
                        rec.Type = line.Substring(1, n - 1);
                        rec.SubType = line.Substring(n + 2).Replace(":", "");

                        var session = this.Sessions[this.Sessions.Count - 1];
                        session.Records.Add(rec);
                        state = LogState.InRecord;
                    }
                    else
                    {
                        throw new System.ArgumentException();
                    }
                }
                else if (state == LogState.InRecord)
                {
                    if (line.Length == 0)
                    {
                        state = LogState.InSession;
                    }
                    else if (line.StartsWith("Context:"))
                    {
                        var session = this.Sessions[this.Sessions.Count - 1];
                        var rec = session.Records[session.Records.Count - 1];
                        rec.Context = line.Substring("Context:".Length);
                    }
                    else if (line.StartsWith("Description:"))
                    {
                        var session = this.Sessions[this.Sessions.Count - 1];
                        var rec = session.Records[session.Records.Count - 1];
                        rec.Description = line.Substring("Description:".Length);
                    }
                    else
                    {
                        throw new System.ArgumentException();
                    }
                }
            }
        }

        private static List<string> GetLinesSharedRead(string filename)
        {
            var lines = new List<string>();
            using (
                var inStream = new System.IO.FileStream(
                    filename,
                    System.IO.FileMode.Open,
                    System.IO.FileAccess.Read,
                    System.IO.FileShare.ReadWrite))
            {
                using (var sr = new System.IO.StreamReader(inStream))
                {
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        lines.Add(line);
                    }
                }
            }
            return lines;
        }
    }
}
using System.Collections.Generic;

namespace VisioAutomation.Application.Logging
{
    public class LogFile
    {
        public string Source;
        public List<LogSession> Sessions;

        public LogFile(string filename)
        {
            this.Sessions = new List<LogSession>();

            var state = LogState.Start;
            var fp = System.IO.File.OpenText(filename);
            string rawline;
            while ((rawline = fp.ReadLine()) != null)
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
                    else if (line.EndsWith("End Session"))
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
    }
}
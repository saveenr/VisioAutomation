using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Application.Logging
{
    public class XmlErrorLog
    {
        public List<FileSessions> FileSessions;

        public XmlErrorLog(string filename)
        {
            this.FileSessions = new List<FileSessions>();

            if (!System.IO.File.Exists(filename))
            {
                string msg = $"File \"{filename}\"does not exist";
                throw new ArgumentException(msg);
            }

            var state = LogState.Start;

            List<string> lines = XmlErrorLog.GetLinesSharedRead(filename);
            lines.Reverse();

            var q = new Stack<string>( lines);
            
            while (q.Count>0)
            {
                string rawline = q.Pop();
                string line = rawline.Trim();

                if (state == LogState.Start)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    if (line.StartsWith("Open") || line.StartsWith("Insert"))
                    {
                        state = this.StartNewSession(line);
                    }
                    else
                    {
                        throw new Exception();
                    }
                }
                else if (state == LogState.InFileSession)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    if (line.StartsWith("Source:"))
                    {
                        var cur_session = this.GetMostRecentSession();
                        cur_session.Source = line.Substring("Source:".Length).Trim();
                    }
                    else if (line.EndsWith("Begin Session"))
                    {
                        var tokens = line.Split();
                        var cur_session = this.GetMostRecentSession();
                        cur_session.StartTimeRaw = string.Join(" ", tokens.Take(5));

                        // Dates are in this format "Sat Jan 10 20:09:12 2015"

                        cur_session.StartTime = DateTime.ParseExact(cur_session.StartTimeRaw, "ddd MMM dd HH:mm:ss yyyy", System.Globalization.CultureInfo.InvariantCulture);

                    }
                    else if (line.EndsWith("End Session"))
                    {
                        state = this.TerminateCurrentSession(line);
                    }
                    else if (line.StartsWith("Open") || line.StartsWith("Insert"))
                    {
                        state = this.TerminateCurrentSession(line);
                        q.Push(line);
                    }
                    else if (line.StartsWith("["))
                    {
                        state = this.StartRecord(line, state);
                    }
                    else
                    {
                        throw new ArgumentException();
                    }
                }
                else if (state == LogState.InRecord)
                {
                    if (line.Length == 0)
                    {
                        state = LogState.InFileSession;
                    }
                    else if (line.StartsWith("Context:"))
                    {
                        var session = this.GetMostRecentSession();
                        var rec = session.Records[session.Records.Count - 1];
                        rec.Context = line.Substring("Context:".Length);
                    }
                    else if (line.StartsWith("Description:"))
                    {
                        var session = this.GetMostRecentSession();
                        var rec = session.Records[session.Records.Count - 1];
                        rec.Description = line.Substring("Description:".Length);
                    }
                    else
                    {
                        throw new ArgumentException();
                    }
                }
                else
                {
                    throw new ArgumentException();                    
                }
            }
        }

        private LogState StartRecord(string line, LogState state)
        {
            var rec = new LogRecord();
            int n = line.IndexOf(']');
            if (n < 2)
            {
                throw new ArgumentException();
            }
            rec.Type = line.Substring(1, n - 1);
            rec.SubType = line.Substring(n + 2).Replace(":", "");

            var session = this.FileSessions[this.FileSessions.Count - 1];
            session.Records.Add(rec);
            state = LogState.InRecord;
            return state;
        }

        private LogState StartNewSession(string line )
        {
            var session = new FileSessions();
            session.StartLine = line;
            var tokens = line.Split();
            session.FileType = tokens[1];

            this.FileSessions.Add(session);

            return LogState.InFileSession;
        }

        private LogState TerminateCurrentSession(string line)
        {
            var session = this.GetMostRecentSession();
            session.EndLine = line;
            return LogState.Start;
        }

        public FileSessions GetMostRecentSession()
        {
            return this.FileSessions[this.FileSessions.Count - 1];
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
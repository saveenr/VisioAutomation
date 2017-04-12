using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Logging
{
    public class XmlErrorLog
    {
        public List<LogSession> LogSessions;

        public XmlErrorLog(string filename)
        {
            this.LogSessions = new List<LogSession>();

            if (!System.IO.File.Exists(filename))
            {
                string msg = string.Format("File \"{0}\"does not exist", filename);
                throw new ArgumentException(msg);
            }

            var state = LogState.Start;

            var lines = XmlErrorLog.GetLines(filename);
            lines.Reverse();

            var stack = new Stack<string>(lines);

            while (stack.Count > 0)
            {
                string rawline = stack.Pop();
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
                        throw new System.ArgumentException("Unexpected Input in LogState.Start");
                    }
                }
                else if (state == LogState.InFileSession)
                {
                    if (line.Length == 0)
                    {
                        continue;
                    }

                    string source = GetStringAfterStartsWith(line, "Source:");

                    if (source != null)
                    {
                        var cur_session = this.GetMostRecentSession();
                        cur_session.Source = source.Trim();
                    }
                    else if (line.EndsWith("Begin Session"))
                    {
                        var tokens = line.Split();
                        var cur_session = this.GetMostRecentSession();
                        cur_session.StartTimeRaw = string.Join(" ", tokens.Take(5));

                        // Dates are in this format "Sat Jan 10 20:09:12 2015"

                        cur_session.StartTime = DateTime.ParseExact(cur_session.StartTimeRaw, "ddd MMM dd HH:mm:ss yyyy",
                            System.Globalization.CultureInfo.InvariantCulture);

                    }
                    else if (line.EndsWith("End Session"))
                    {
                        state = this.TerminateCurrentSession(line);
                    }
                    else if (line.StartsWith("Open") || line.StartsWith("Insert"))
                    {
                        state = this.TerminateCurrentSession(line);
                        stack.Push(line);
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
                    else
                    {
                        string context = GetStringAfterStartsWith(line, "Context:");
                        string description = GetStringAfterStartsWith(line, "Description:");

                        if (context != null)
                        {
                            // Store a Context Record 
                            var session = this.GetMostRecentSession();
                            var rec = session.LogRecords[session.LogRecords.Count - 1];
                            rec.Context = context;
                        }
                        else if (description != null)
                        {
                            // Store a Description Record 
                            var session = this.GetMostRecentSession();
                            var rec = session.LogRecords[session.LogRecords.Count - 1];
                            rec.Description = description;
                        }
                        else
                        {
                            throw new ArgumentException();
                        }
                    }
                }
                else
                {
                    throw new ArgumentException();
                }
            }
        }

        private static string GetStringAfterStartsWith(string line, string text_context)
        {
            if (line.StartsWith(text_context))
            {
                string result = line.Substring(text_context.Length);
                return result;
            }
            return null;
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
            rec.SubType = line.Substring(n + 2).Replace(":", string.Empty);

            var session = this.LogSessions[this.LogSessions.Count - 1];
            session.LogRecords.Add(rec);
            state = LogState.InRecord;
            return state;
        }

        private LogState StartNewSession(string line)
        {
            var session = new LogSession();
            session.StartLine = line;
            var tokens = line.Split();
            session.FileType = tokens[1];

            this.LogSessions.Add(session);

            return LogState.InFileSession;
        }

        private LogState TerminateCurrentSession(string line)
        {
            var session = this.GetMostRecentSession();
            session.EndLine = line;
            return LogState.Start;
        }

        public LogSession GetMostRecentSession()
        {
            return this.LogSessions[this.LogSessions.Count - 1];
        }

        private static List<string> GetLines(string filename)
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
using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Application.Logging
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

            var lines = XmlErrorLog._get_lines(filename);
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
                        state = this._start_new_session(line);
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

                    string source = _get_string_after_starts_with(line, "Source:");

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

                        var culture = System.Globalization.CultureInfo.InvariantCulture;
                        cur_session.StartTime = DateTime.ParseExact(cur_session.StartTimeRaw, "ddd MMM dd HH:mm:ss yyyy",
                            culture);

                    }
                    else if (line.EndsWith("End Session"))
                    {
                        state = this._terminate_current_session(line);
                    }
                    else if (line.StartsWith("Open") || line.StartsWith("Insert"))
                    {
                        state = this._terminate_current_session(line);
                        stack.Push(line);
                    }
                    else if (line.StartsWith("["))
                    {
                        state = this._start_record(line, state);
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
                        string context = _get_string_after_starts_with(line, "Context:");
                        string description = _get_string_after_starts_with(line, "Description:");

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

        private static string _get_string_after_starts_with(string line, string text_context)
        {
            if (line.StartsWith(text_context))
            {
                string result = line.Substring(text_context.Length);
                return result;
            }
            return null;
        }

        private LogState _start_record(string line, LogState state)
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

        private LogState _start_new_session(string line)
        {
            var session = new LogSession();
            session.StartLine = line;
            var tokens = line.Split();
            session.FileType = tokens[1];

            this.LogSessions.Add(session);

            return LogState.InFileSession;
        }

        private LogState _terminate_current_session(string line)
        {
            var session = this.GetMostRecentSession();
            session.EndLine = line;
            return LogState.Start;
        }

        public LogSession GetMostRecentSession()
        {
            return this.LogSessions[this.LogSessions.Count - 1];
        }

        private static List<string> _get_lines(string filename)
        {
            var lines = new List<string>();
            using (
                var in_stream = new System.IO.FileStream(
                    filename,
                    System.IO.FileMode.Open,
                    System.IO.FileAccess.Read,
                    System.IO.FileShare.ReadWrite))
            {
                using (var sr = new System.IO.StreamReader(in_stream))
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
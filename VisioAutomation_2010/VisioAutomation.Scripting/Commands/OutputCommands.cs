using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class OutputCommands : CommandSet
    {
        public OutputCommands(Session session) :
            base(session)
        {

        }

        private void Write(OutputStream output, string s)
        {
            this.Session.Write(output,s);
        }

        private void Write(OutputStream output, string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this.Session.Write(output, s);
        }

        public void WriteUser(string s)
        {
            this.Write(OutputStream.User, s);
        }

        public void WriteDebug(string s)
        {
            this.Write(OutputStream.Debug, s);
        }

        public void WriteError(string s)
        {
            this.Write(OutputStream.Error, s);
        }

        public void WriteVerbose(string s)
        {
            this.Write(OutputStream.Verbose, s);
        }

        public void WriteUser(string fmt, params object[] items)
        {
            this.Write(OutputStream.User, fmt, items);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            this.Write(OutputStream.Debug,fmt,items);           
        }

        public void WriteError(string fmt, params object[] items)
        {
            this.Write(OutputStream.Error, fmt, items);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            this.Write(OutputStream.Verbose, fmt, items);
        }
    }
}
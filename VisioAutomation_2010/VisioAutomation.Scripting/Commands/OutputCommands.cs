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

        public void WriteUser(string s)
        {
            this.Session.WriteUser(s);
        }

        public void WriteDebug(string s)
        {
            this.Session.WriteDebug(s);
        }

        public void WriteError(string s)
        {
            this.Session.WriteError(s);
        }

        public void WriteVerbose(string s)
        {
            this.Session.WriteVerbose(s);
        }

        public void WriteUser(string fmt, params object[] items)
        {
            this.Session.WriteUser(fmt, items);
        }

        public void WriteDebug(string fmt, params object[] items)
        {
            this.Session.WriteDebug( fmt, items);           
        }

        public void WriteError(string fmt, params object[] items)
        {
            this.Session.WriteError( fmt, items);
        }

        public void WriteVerbose(string fmt, params object[] items)
        {
            this.Session.WriteVerbose( fmt, items);
        }
    }
}
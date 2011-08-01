using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class OutputCommands : SessionCommands
    {
        public OutputCommands(Session session) :
            base(session)
        {

        }

        internal void Write(OutputStream output, string s)
        {
            this.Session.Write(output,s);
        }

        public void Write(OutputStream output, string fmt, params object[] items)
        {
            string s = string.Format(fmt, items);
            this.Session.Write(output, s);
        }
    }
}
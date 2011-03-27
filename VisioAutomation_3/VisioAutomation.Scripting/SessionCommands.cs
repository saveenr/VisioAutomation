using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class SessionCommands
    {
        // Keep a reference back to the parent session. This gives access to all other commands
        // for a the current context
        protected readonly Session Session;

        public SessionCommands(Session session)
        {
            this.Session = session;
        }       
    }
}
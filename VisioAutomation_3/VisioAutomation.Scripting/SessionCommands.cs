using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class SessionCommands
    {
        protected readonly Session Session;

        public SessionCommands(Session session)
        {
            this.Session = session;
        }       
    }
}
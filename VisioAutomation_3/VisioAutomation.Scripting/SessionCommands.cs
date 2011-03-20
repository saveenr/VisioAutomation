using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting
{
    public class SessionCommands
    {
        public readonly Session Session;

        public SessionCommands(Session session)
        {
            this.Session = session;
        }

        public bool HasSelectedShapes()
        {
            return this.Session.HasSelectedShapes();
        }

        public bool HasSelectedShapes(int min_items)
        {
            return this.Session.HasSelectedShapes(min_items);
        }

        public bool HasActiveDrawing()
        {
            return this.Session.HasActiveDrawing();
        }

        public IVisio.Application Application
        {
            get { return this.Session.Application; }
        }
        
    }
}
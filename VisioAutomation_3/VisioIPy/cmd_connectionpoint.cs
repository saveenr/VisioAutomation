using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void AddConnectionPoint(string x, string y, VA.Connections.ConnectionPointType type)
        {
            this.ScriptingSession.ConnectionPoint.AddConnectionPoint(x, y, type);
        }

        public void DeleteConnectionPoint(int index)
        {
            this.ScriptingSession.ConnectionPoint.DeleteConnectionPoint(index);
        }

        public IDictionary<IVisio.Shape, IList<VA.Connections.ConnectionPointCells>> GetConnectionPoints()
        {
            return this.ScriptingSession.ConnectionPoint.GetConnectionPoints();
        }
    }
}
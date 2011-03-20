using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Shape ConnectShapes(IVisio.Shape from, IVisio.Shape to)
        {
            this.ScriptingSession.Document.OpenStencil("basic_u.vss");
            var master = this.ScriptingSession.Master.GetMaster("Dynamic Connector", "basic_u.vss");

            var fromshapes = new [] {from};
            var toshapes = new [] {to};
            var connectors = this.ScriptingSession.Connection.ConnectShapes(master, fromshapes, toshapes);
            return connectors[0];
        }
    }
}
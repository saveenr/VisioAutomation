using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Shape DropMaster(IVisio.Master master, double x, double y)
        {
            var ss = this.ScriptingSession;
            return ss.Master.DropMaster(master, x, y);
        }

        public short[] DropMasters(IList<IVisio.Master> masters, IList<VA.Drawing.Point> points)
        {
            var ss = this.ScriptingSession;
            return ss.Master.DropMasters(masters, points);
        }

        public IVisio.Master GetMaster(string master, string stencil)
        {
            var ss = this.ScriptingSession;
            var m = ss.Master.GetMaster(master, stencil);
            return m;
        }

        public IList<IVisio.Master> GetMasters()
        {
            var ss = this.ScriptingSession;
            return ss.Master.GetMasters();
        }

        public IVisio.Master NewMaster(IVisio.Document stencil, string name)
        {
            var master = ScriptingSession.Master.NewMaster(stencil, name);
            return master;
        }
    }
}
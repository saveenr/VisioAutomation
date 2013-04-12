using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class MasterCommands : CommandSet
    {
        public MasterCommands(Session session) :
            base(session)
        {

        }

        public IList<IVisio.Master> Get()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IList<IVisio.Master> Get(IVisio.Document doc)
        {
            this.CheckVisioApplicationAvailable();
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IVisio.Master Get(string name)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("mastername");
            }


            IVisio.Master master = null;
            try
            {
                var application = this.Session.VisioApplication;
                var active_document = application.ActiveDocument;
                var masters = active_document.Masters;
                master = masters.ItemU[name];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No such master \"{0}\"", name);
                throw new VA.Scripting.ScriptingException(msg);
            }
            return master;
        }

        public IVisio.Master Get(string master, IVisio.Document stencil)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            if (stencil == null)
            {
                throw new System.ArgumentNullException("stencil");
            }

            var application = this.Session.VisioApplication;
            var documents = application.Documents;

            var masters = stencil.Masters;
            IVisio.Master masterobj = this.TryGetMaster(masters, master);
            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", master, stencil);
                throw new VA.Scripting.ScriptingException(msg);
            }

            return masterobj;
        }

        private IVisio.Master TryGetMaster(IVisio.Masters masters, string name)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            try
            {
                var masterobj = masters.ItemU[name];
                return masterobj;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        public IVisio.Shape Drop(IVisio.Master master, double x, double y)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shape = page.Drop(master, x, y);
            return shape;
        }

        public short[] Drop(IList<IVisio.Master> masters, IList<VA.Drawing.Point> points)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            if (masters == null)
            {
                throw new System.ArgumentNullException("points");
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master New(IVisio.Document stencil, string name)
        {
            this.CheckVisioApplicationAvailable();

            var masters = stencil.Masters;
            var master = masters.AddEx(IVisio.VisMasterTypes.visTypeMaster);
            master.Name = name;

            return master;
        }
    }
}
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
            if (!this.Session.HasActiveDrawing)
            {
                this.Session.Write(OutputStream.Verbose,"No Active Document - 0 Masters");
                new List<IVisio.Master>(0);
            }

            var application = this.Session.VisioApplication;
            var doc = application.ActiveDocument;
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IList<IVisio.Master> Get(IVisio.Document doc)
        {
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IVisio.Master Get(string name)
        {
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
                throw new AutomationException(msg);
            }
            return master;
        }

        private static IVisio.Document TryGetDocument(IVisio.Documents documents, string name)
        {
            try
            {
                var stencil_doc = documents[name];
                return stencil_doc;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        public IVisio.Master Get(string master, string stencil)
        {
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
            IVisio.Document stencil_doc = TryGetDocument(documents,stencil);
            if (stencil_doc==null)
            {
                string msg = string.Format("No such stencil \"{0}\"", stencil);
                throw new AutomationException(msg);
            }

            var masters = stencil_doc.Masters;
            IVisio.Master masterobj = this.TryGetMaster(masters,master);
            if (masterobj==null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", master, stencil);
                throw new AutomationException(msg);
            }

            return masterobj;
        }

        private IVisio.Master TryGetMaster(IVisio.Masters masters, string name)
        {
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
            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shape = page.Drop(master, x, y);
            return shape;
        }

        public short[] Drop(IList<IVisio.Master> masters, IList<VA.Drawing.Point> points)
        {
            if (masters == null)
            {
                throw new System.ArgumentNullException("points");
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            if (!this.Session.HasActiveDrawing)
            {
                throw new AutomationException("No active page");
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master New(IVisio.Document stencil, string name)
        {

            var masters = stencil.Masters;

            var master = masters.AddEx(IVisio.VisMasterTypes.visTypeMaster);
            master.Name = name;

            return master;
        }
    }
}
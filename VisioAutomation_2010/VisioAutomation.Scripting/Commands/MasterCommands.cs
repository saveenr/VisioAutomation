using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class MasterCommands : CommandSet
    {
        internal MasterCommands(Client client) :
            base(client)
        {

        }

        public void OpenForEdit(IVisio.Master master)
        {
            var mdraw_window = master.OpenDrawWindow();
            mdraw_window.Activate();
        }

        public void CloseMasterEditing()
        {
            var window = this.Client.VisioApplication.ActiveWindow;

            var st = window.SubType;
            if (st != 64)
            {
                throw new AutomationException("The active window is not a master window");
            }

            var master = (IVisio.Master)window.Master;
            master.Close();
        }

        public IList<IVisio.Master> Get()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var doc = application.ActiveDocument;
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IList<IVisio.Master> Get(IVisio.Document doc)
        {
            this.Client.Application.AssertApplicationAvailable();
            var doc_masters = doc.Masters;
            var masters = doc_masters.AsEnumerable().ToList();
            return masters;
        }

        public IVisio.Master Get(string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("mastername");
            }

            IVisio.Master master;
            try
            {
                var application = this.Client.VisioApplication;
                var active_document = application.ActiveDocument;
                var masters = active_document.Masters;
                master = masters.ItemU[name];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                string msg = string.Format("No such master \"{0}\"", name);
                throw new VisioOperationException(msg);
            }
            return master;
        }

        public IVisio.Master Get(string master, IVisio.Document doc)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (master == null)
            {
                throw new System.ArgumentNullException("master");
            }

            if (doc == null)
            {
                throw new System.ArgumentNullException("doc");
            }

            var application = this.Client.VisioApplication;
            var documents = application.Documents;

            var masters = doc.Masters;
            IVisio.Master masterobj = this.TryGetMaster(masters, master);
            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", master, doc);
                throw new VisioOperationException(msg);
            }

            return masterobj;
        }

        public List<IVisio.Master> GetMastersByName(string name, IVisio.Document doc)
        {
            if (name == null || name == "*")
            {
                // return all masters
                var masters = doc.Masters.AsEnumerable().ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = doc.Masters.AsEnumerable();
                var masters3 = TextUtil.FilterObjectsByNames(masters2, new[] { name }, p => p.Name, true, TextUtil.FilterAction.Include).ToList();
                return masters3;
            }
        }

        public List<IVisio.Master> GetMastersByName(string name)
        {
            var application = this.Client.VisioApplication;
            var doc = application.ActiveDocument;
            return this.GetMastersByName(name, doc);
        }

        private IVisio.Master TryGetMaster(IVisio.Masters masters, string name)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

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
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            var shape = page.Drop(master, x, y);
            return shape;
        }

        public short[] Drop(IList<IVisio.Master> masters, IList<Drawing.Point> points)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (masters == null)
            {
                throw new System.ArgumentNullException("points");
            }

            if (points == null)
            {
                throw new System.ArgumentNullException("points");
            }

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master New(IVisio.Document document, string name)
        {
            this.Client.Application.AssertApplicationAvailable();

            if (document == null)
            {
                document = this.Client.VisioApplication.ActiveDocument;
                if (document == null)
                {
                    throw new AutomationException("No Active Document");
                }
            }

            var masters = document.Masters;
            var master = masters.AddEx(IVisio.VisMasterTypes.visTypeMaster);
            if (name != null)
            {
                master.Name = name;
            }

            return master;
        }

        // http://blogs.msdn.com/b/visio/archive/2010/01/27/container-list-and-callout-api-in-visio-2010.aspx
        // https://msdn.microsoft.com/en-us/library/office/ff768907(v=office.14).aspx
        public IVisio.Shape DropContainer(IVisio.Master master)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            var selectedShapes = this.Client.Selection.Get();

            var shape = page.DropContainer(master, selectedShapes);
            return shape;
        }

    }
}
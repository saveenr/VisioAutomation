using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
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
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            var application = cmdtarget.Application;
            var window = application.ActiveWindow;

            var st = window.SubType;
            if (st != 64)
            {
                throw new System.ArgumentException("The active window is not a master window");
            }

            var master = (IVisio.Master)window.Master;
            master.Close();
        }

        public List<IVisio.Master> Get()
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            var application = cmdtarget.Application;
            var doc = application.ActiveDocument;
            var doc_masters = doc.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public List<IVisio.Master> Get(IVisio.Document doc)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            var doc_masters = doc.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public IVisio.Master Get(string name)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("mastername");
            }

            IVisio.Master master;
            try
            {
                var application = cmdtarget.Application;
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
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            if (doc == null)
            {
                throw new System.ArgumentNullException(nameof(doc));
            }

            var application = cmdtarget.Application;
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
                var masters = doc.Masters.ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = doc.Masters.ToEnumerable();
                var masters3 = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(masters2, new[] { name }, p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                return masters3;
            }
        }

        public List<IVisio.Master> GetMastersByName(string name)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);

            var application = cmdtarget.Application;
            var doc = application.ActiveDocument;
            return this.GetMastersByName(name, doc);
        }

        private IVisio.Master TryGetMaster(IVisio.Masters masters, string name)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


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

        public IVisio.Shape Drop(IVisio.Master master, VisioAutomation.Geometry.Point p)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var page = application.ActivePage;
            var shape = page.Drop(master, p.X, p.Y);
            return shape;
        }

        public short[] Drop(IList<IVisio.Master> masters, IList<VisioAutomation.Geometry.Point> points)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            var application = cmdtarget.Application;
            var page = application.ActivePage;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master New(IVisio.Document document, string name)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application);


            if (document == null)
            {
                var application = cmdtarget.Application;
                document = application.ActiveDocument;
                if (document == null)
                {
                    throw new System.ArgumentException("No Active Document");
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
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var page = application.ActivePage;
            var selectedShapes = this._client.Selection.Get();

            var shape = page.DropContainer(master, selectedShapes);
            return shape;
        }

        public IVisio.Shape DropContainer(string master)
        {
            var cmdtarget = new CommandTarget(this._client, CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            var application = cmdtarget.Application;
            var page = application.ActivePage;
            var selectedShapes = this._client.Selection.Get();

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = application.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = application.Documents.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[master];
            var shape = page.DropContainer(container_master,selectedShapes);

            return shape;
        }

    }
}
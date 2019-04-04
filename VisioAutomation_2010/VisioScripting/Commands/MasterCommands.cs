using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using VisioScripting.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class MasterCommands : CommandSet
    {
        internal MasterCommands(Client client) :
            base(client)
        {

        }

        public void OpenMasterForEdit(IVisio.Master master)
        {
            var mdraw_window = master.OpenDrawWindow();
            mdraw_window.Activate();
        }

        public void CloseMasterEditing()
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            var window = cmdtarget.Application.ActiveWindow;

            var win_subtype = window.SubType;
            if (win_subtype != 64)
            {
                throw new System.ArgumentException("The active window is not a master window");
            }

            var master = (IVisio.Master)window.Master;
            master.Close();
        }

        public List<IVisio.Master> GetMasters(TargetDocument targetdoc)
        {
            var doc_masters = targetdoc.Item.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public IVisio.Master GetMaster(TargetDocument targetdoc, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var masters = targetdoc.Item.Masters;
            IVisio.Master masterobj = this.TryGetMaster(masters, name);

            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", name, targetdoc.Item.Name);
                throw new VisioOperationException(msg);
            }

            return masterobj;
        }


        public List<IVisio.Master> FindMasters(TargetDocument targetdoc, string name)
        {
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all masters
                var masters = targetdoc.Item.Masters.ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = targetdoc.Item.Masters.ToEnumerable();
                var masters3 = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(masters2, new[] { name }, p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                return masters3;
            }
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

        public IVisio.Shape DropMaster(TargetPage targetpage, IVisio.Master master, VisioAutomation.Geometry.Point p)
        {
            targetpage = targetpage.Resolve(this._client);

            var shape = targetpage.Item.Drop(master, p.X, p.Y);
            return shape;
        }

        public short[] DropMasters(
            TargetPage targetpage,
            IList<IVisio.Master> masters, 
            IList<VisioAutomation.Geometry.Point> points)
        {
            targetpage = targetpage.Resolve(this._client);

            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            var page = targetpage.Item;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        public IVisio.Master NewMaster(IVisio.Document document, string name)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (document == null)
            {
                document = cmdtarget.ActiveDocument;
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

        public IVisio.Shape DropContainerMaster(TargetPage targetpage, IVisio.Master master)
        {
            var page = targetpage.Item;
            var app = page.Application;
            var window = app.ActiveWindow;
            var selection = window.Selection;

            var shape = page.DropContainer(master, selection);
            return shape;
        }

        public IVisio.Shape DropContainer(TargetPage targetpage, string master)
        {
            var page = targetpage.Item;
            var app = page.Application;
            var window = app.ActiveWindow;
            var selection = window.Selection;
            var docs = app.Documents;

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = app.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = docs.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[master];
            var shape = page.DropContainer(container_master,selection);

            return shape;
        }
    }
}
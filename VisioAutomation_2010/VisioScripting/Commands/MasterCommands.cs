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

        public List<IVisio.Master> GetAllMastersInDocument(TargetDocument target_doc)
        {
            var doc_masters = target_doc.Item.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public IVisio.Master GetMasterWithNameInDocument(TargetDocument target_doc, string master)
        {
            if (master == null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            var masters = target_doc.Item.Masters;
            IVisio.Master masterobj = this.TryGetMaster(masters, master);

            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", master, target_doc.Item.Name);
                throw new VisioOperationException(msg);
            }

            return masterobj;
        }


        public List<IVisio.Master> FindMastersInDocumentByName(TargetDocument target_doc, string name)
        {
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all masters
                var masters = target_doc.Item.Masters.ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = target_doc.Item.Masters.ToEnumerable();
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

        public IVisio.Shape DropMasterOnActivePage(IVisio.Master master, VisioAutomation.Geometry.Point p)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            var shape = page.Drop(master, p.X, p.Y);
            return shape;
        }

        public short[] DropMastersOnActivePage(
            IList<IVisio.Master> masters, 
            IList<VisioAutomation.Geometry.Point> points)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            if (masters == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            if (points == null)
            {
                throw new System.ArgumentNullException(nameof(points));
            }

            var page = cmdtarget.ActivePage;
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

        public IVisio.Shape DropContainerMaster(IVisio.Master master)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;

            var shape = page.DropContainer(master, selection);
            return shape;
        }

        public IVisio.Shape DropContainer(string master)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();


            var application = cmdtarget.Application;
            var page = application.ActivePage;
            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = application.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = application.Documents.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[master];
            var shape = page.DropContainer(container_master,selection);

            return shape;
        }
    }
}
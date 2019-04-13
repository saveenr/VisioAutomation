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

        public List<IVisio.Master> GetMasters(TargetDocument targetdoc)
        {
            var doc_masters = targetdoc.Document.Masters;
            var masters = doc_masters.ToList();
            return masters;
        }

        public IVisio.Master GetMaster(TargetDocument targetdoc, string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            var masters = targetdoc.Document.Masters;
            IVisio.Master masterobj = this._try_get_master(masters, name);

            if (masterobj == null)
            {
                string msg = string.Format("No such master \"{0}\" in \"{1}\"", name, targetdoc.Document.Name);
                throw new VisioOperationException(msg);
            }

            return masterobj;
        }


        public List<IVisio.Master> FindMasters(TargetDocument targetdoc, string name)
        {
            if (VisioScripting.Helpers.WildcardHelper.NullOrStar(name))
            {
                // return all masters
                var masters = targetdoc.Document.Masters.ToList();
                return masters;
            }
            else
            {
                // return masters matching the name
                var masters2 = targetdoc.Document.Masters.ToEnumerable();
                var masters3 = VisioScripting.Helpers.WildcardHelper.FilterObjectsByNames(masters2, new[] { name }, p => p.Name, true, VisioScripting.Helpers.WildcardHelper.FilterAction.Include).ToList();
                return masters3;
            }
        }

        private IVisio.Master _try_get_master(IVisio.Masters masters, string name)
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

            var shape = targetpage.Page.Drop(master, p.X, p.Y);
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

            var page = targetpage.Page;
            var shapeids = page.DropManyU(masters, points);
            return shapeids;
        }

        // http://blogs.msdn.com/b/visio/archive/2010/01/27/container-list-and-callout-api-in-visio-2010.aspx
        // https://msdn.microsoft.com/en-us/library/office/ff768907(v=office.14).aspx

        public IVisio.Shape DropContainerMaster(TargetPage targetpage, IVisio.Master master)
        {
            targetpage = targetpage.Resolve(this._client);
            var page = targetpage.Page;
            var app = page.Application;
            var window = app.ActiveWindow;
            var selection = window.Selection;

            var shape = page.DropContainer(master, selection);
            return shape;
        }

        public IVisio.Shape DropContainer(TargetPage targetpage, string master)
        {
            targetpage = targetpage.Resolve(this._client);
            var page = targetpage.Page;
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
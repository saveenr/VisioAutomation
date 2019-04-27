
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class ContainerCommands : CommandSet
    {
        internal ContainerCommands(Client client) :
            base(client)
        {

        }

        public IVisio.Shape DropContainerMaster(TargetPage targetpage, IVisio.Master master)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            var page = targetpage.Page;
            var app = page.Application;
            var window = app.ActiveWindow;
            var selection = window.Selection;

            var shape = page.DropContainer(master, selection);
            return shape;
        }

        public IVisio.Shape DropContainer(TargetPage targetpage, string master)
        {
            targetpage = targetpage.ResolveToPage(this._client);
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
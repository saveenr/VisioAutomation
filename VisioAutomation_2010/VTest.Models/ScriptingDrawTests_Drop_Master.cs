using System.Linq;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Scripting
{
    public partial class ScriptingDrawTests : Framework.VTest
    {

      
        [MUT.TestMethod]
        public void Scripting_Drop_Master()
        {
            var pagesize = new VA.Core.Size(4, 4);
            var client = this.GetScriptingClient();

            // Create the page
            client.Document.NewDocument();

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Frop the Shapes

            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Core.Point(2, 2));

            // Verify
            var application = client.Application.GetApplication();
            var active_page = application.ActivePage;
            var shapes = active_page.Shapes;
            MUT.Assert.AreEqual(1, shapes.Count);

            // cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Drop_Many()
        {
            var pagesize = new VA.Core.Size(10, 10);
            var client = this.GetScriptingClient();

            // Create the Page
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var m1 = client.Master.GetMaster(stencil_targetdoc, "Rectangle");
            var m2 = client.Master.GetMaster(stencil_targetdoc, "Ellipse");

            // Drop the Shapes
            var masters = new[] {m1, m2};
            var xys = new[] { 1.0, 2.0, 3.0, 4.0, 1.5, 4.5, 5.7, 2.4 };
            var points = VA.Core.Point.FromDoubles(xys).ToList();

            client.Master.DropMasters(VisioScripting.TargetPage.Auto, masters, points);

            // Verify
            var application = client.Application.GetApplication();
            MUT.Assert.AreEqual(4, application.ActivePage.Shapes.Count);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Drop_Container_Master_Object()
        {
            var pagesize = new VA.Core.Size(4, 4);
            var client = this.GetScriptingClient();


            // Create the page
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var application = client.Application.GetApplication();
            var active_page = application.ActivePage;

            // Load the stencils and find the masters
            var basic_stencil = client.Document.OpenStencilDocument("Basic_U.VSS");
            var stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(stencil_targetdoc, "Rectangle");

            // Drop the rectangle
            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Core.Point(2, 2) );

            // Select the rectangle... it should already be selected, but just make sure


            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var app = active_page.Application;

            var ver = client.Application.ApplicationVersion;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";

            var stencil_type = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;
            var measurement_system = IVisio.VisMeasurementSystem.visMSUS;
            var containers_file = app.GetBuiltInStencilFile(stencil_type, measurement_system);
            var containers_doc = app.Documents.OpenStencil(containers_file);
            var masters = containers_doc.Masters;
            var container_master = masters.ItemU[cont_master_name];

            var dropped_container = client.Container.DropContainerMaster(VisioScripting.TargetPage.Auto, container_master);

            // Verify
            var shapes = active_page.Shapes;
            // There should be two shapes... the rectangle and the container
            MUT.Assert.AreEqual(2, shapes.Count);

            // Verify that we did indeed drop a container

            var results_dic = VisioAutomation.Shapes.UserDefinedCellHelper.GetDictionary(dropped_container, VA.Core.CellValueType.Result);
            MUT.Assert.IsTrue(results_dic.ContainsKey("msvStructureType"));
            var prop = results_dic["msvStructureType"];
            MUT.Assert.AreEqual("Container", prop.Value.Value);

            // cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

    }
}
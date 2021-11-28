using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Models
{
    public partial class ScriptingDrawTests : Framework.VTest
    {



        [MUT.TestMethod]
        public void Scripting_Drop_Container_Master_Name()
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
            var basic_stencil_targetdoc = new VisioScripting.TargetDocument(basic_stencil);
            var master = client.Master.GetMaster(basic_stencil_targetdoc, "Rectangle");

            // Drop the rectangle
            client.Master.DropMaster(VisioScripting.TargetPage.Auto, master, new VA.Core.Point(2, 2) );


            // Select the rectangle... it should already be selected, but just make sure
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            // Drop the container... since the rectangle is selected... it will automatically make it a member of the container
            var ver = client.Application.ApplicationVersion;
            var cont_master_name = ver.Major >= 15 ? "Plain" : "Container 1";
            var dropped_container = client.Container.DropContainer(VisioScripting.TargetPage.Auto, cont_master_name);

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
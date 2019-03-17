using Microsoft.VisualStudio.TestTools.UnitTesting;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models.Layouts
{
    [TestClass]
    public class DirectedGraph_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void DirectedGraph_WithBezierConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();
            
            var options = new VADG.MsaglLayoutOptions();
            options.UseDynamicConnectors = false;
            
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;
            directed_graph_drawing.Render(page,options);

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(DirectedGraph_WithBezierConnectors),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [TestMethod]
        public void DirectedGraph_WithDynamicConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();

            var options = new VADG.MsaglLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;
            
            directed_graph_drawing.Render(page1,options);

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(DirectedGraph_WithDynamicConnectors),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [TestMethod]
        public void RenderDirectedGraphWithCustomProps()
        {
            var d = new VADG.DirectedGraphLayout();

            var n0 = d.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Geometry.Size(3, 2);
            n0.CustomProperties = new CustomPropertyDictionary();
            n0.CustomProperties["p1"] = new CustomPropertyCells("\"v1\"");
            n0.CustomProperties["p2"] = new CustomPropertyCells("\"v2\"");
            n0.CustomProperties["p3"] = new CustomPropertyCells("\"v3\"");

            var options = new VADG.MsaglLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            d.Render(page1, options);
            
            Assert.IsNotNull(n0.VisioShape);
            var props_dic = CustomPropertyHelper.GetCells(n0.VisioShape, CellValueType.Formula);
            Assert.IsTrue(props_dic.Count>=3);
            Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Value);
            Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Value);
            Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Value);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(RenderDirectedGraphWithCustomProps),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        private VADG.DirectedGraphLayout create_sample_graph()
        {
            var d = new VADG.DirectedGraphLayout();

            var basic_stencil = "basic_u.vss";
            var n0 = d.AddShape("n0", "Node 0", basic_stencil, "Rectangle");
            n0.Size = new VA.Geometry.Size(3, 2);
            var n1 = d.AddShape("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = d.AddShape("n2", "Node 2", basic_stencil, "Rectangle");
            var n3 = d.AddShape("n3", "Node 3", basic_stencil, "Rectangle");
            var n4 = d.AddShape("n4", "Node 4\nUnconnected", basic_stencil, "Rectangle");

            var c0 = d.AddConnection("c0", n0, n1, "0 -> 1", VisioAutomation.Models.ConnectorType.Curved);
            var c1 = d.AddConnection("c1", n1, n2, "1 -> 2", VisioAutomation.Models.ConnectorType.RightAngle);
            var c2 = d.AddConnection("c2", n1, n0, "0 -> 1", VisioAutomation.Models.ConnectorType.Curved);
            var c3 = d.AddConnection("c3", n0, n2, "0 -> 2", VisioAutomation.Models.ConnectorType.Straight);
            var c4 = d.AddConnection("c4", n2, n3, "2 -> 3", VisioAutomation.Models.ConnectorType.Curved);
            var c5 = d.AddConnection("c5", n3, n0, "3 -> 0", VisioAutomation.Models.ConnectorType.Curved);

            return d;
        }
    }
}
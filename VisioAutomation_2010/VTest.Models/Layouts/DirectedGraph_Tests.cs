using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using VisioAutomation.Shapes;
using VTest.Framework;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VTest.Models.Layouts
{
    [MUT.TestClass]
    public class DirectedGraph_Tests : Framework.VTest
    {
        [MUT.TestMethod]
        public void DirectedGraph_WithBezierConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();
            
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;


            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = false;
            renderer.Render(page, directed_graph_drawing);
            
            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(DirectedGraph_WithBezierConnectors),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [MUT.TestMethod]
        public void DirectedGraph_WithDynamicConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = true;
            renderer.Render(page1, directed_graph_drawing);

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(DirectedGraph_WithDynamicConnectors),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [MUT.TestMethod]
        public void RenderDirectedGraphWithCustomProps()
        {
            var d = new VADG.DirectedGraphLayout();

            var n0 = d.AddNode("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Core.Size(3, 2);
            n0.CustomProperties = new CustomPropertyDictionary();
            n0.CustomProperties["p1"] = new CustomPropertyCells("\"v1\"");
            n0.CustomProperties["p2"] = new CustomPropertyCells("\"v2\"");
            n0.CustomProperties["p3"] = new CustomPropertyCells("\"v3\"");

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = true;
            renderer.Render(page1, d);

            MUT.Assert.IsNotNull(n0.VisioShape);
            var props_dic = CustomPropertyHelper.GetDictionary(n0.VisioShape, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.IsTrue(props_dic.Count>=3);
            MUT.Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Value);
            MUT.Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Value);
            MUT.Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Value);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(RenderDirectedGraphWithCustomProps),".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        private VADG.DirectedGraphLayout create_sample_graph()
        {
            var d = new VADG.DirectedGraphLayout();

            var basic_stencil = "basic_u.vss";
            var n0 = d.AddNode("n0", "Node 0", basic_stencil, "Rectangle");
            n0.Size = new VA.Core.Size(3, 2);
            var n1 = d.AddNode("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = d.AddNode("n2", "Node 2", basic_stencil, "Rectangle");
            var n3 = d.AddNode("n3", "Node 3", basic_stencil, "Rectangle");
            var n4 = d.AddNode("n4", "Node 4\nUnconnected", basic_stencil, "Rectangle");

            var c0 = d.AddEdge("c0", n0, n1, "0 -> 1", VisioAutomation.Models.ConnectorType.Curved);
            var c1 = d.AddEdge("c1", n1, n2, "1 -> 2", VisioAutomation.Models.ConnectorType.RightAngle);
            var c2 = d.AddEdge("c2", n1, n0, "0 -> 1", VisioAutomation.Models.ConnectorType.Curved);
            var c3 = d.AddEdge("c3", n0, n2, "0 -> 2", VisioAutomation.Models.ConnectorType.Straight);
            var c4 = d.AddEdge("c4", n2, n3, "2 -> 3", VisioAutomation.Models.ConnectorType.Curved);
            var c5 = d.AddEdge("c5", n3, n0, "3 -> 0", VisioAutomation.Models.ConnectorType.Curved);

            return d;
        }
    }
}
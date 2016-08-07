using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VACONNECT = VisioAutomation.Shapes.Connections;
using VACUSTPROP=VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using DG = VisioAutomation.Models.Layouts.DirectedGraph;

namespace VisioAutomation_Tests.Models.Layouts
{
    [TestClass]
    public class DirectedGraph_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void DirectedGraph_WithBezierConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();
            
            var options = new DG.MsaglLayoutOptions();
            options.UseDynamicConnectors = false;
            
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;
            directed_graph_drawing.Render(page,options);

            string output_filename = TestGlobals.TestHelper.GetTestMethodOutputFilename(".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [TestMethod]
        public void DirectedGraph_WithDynamicConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();

            var options = new DG.MsaglLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;
            
            directed_graph_drawing.Render(page1,options);

            string output_filename = TestGlobals.TestHelper.GetTestMethodOutputFilename(".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        [TestMethod]
        public void RenderDirectedGraphWithCustomProps()
        {
            var d = new DG.Drawing();

            var n0 = d.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            n0.CustomProperties = new Dictionary<string, VACUSTPROP.CustomPropertyCells>();
            n0.CustomProperties["p1"] = new VACUSTPROP.CustomPropertyCells("v1");
            n0.CustomProperties["p2"] = new VACUSTPROP.CustomPropertyCells("v2");
            n0.CustomProperties["p3"] = new VACUSTPROP.CustomPropertyCells("v3");

            var options = new DG.MsaglLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            d.Render(page1, options);
            
            Assert.IsNotNull(n0.VisioShape);
            var props_dic = VACUSTPROP.CustomPropertyHelper.Get(n0.VisioShape);
            Assert.IsTrue(props_dic.Count>=3);
            Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Formula);
            Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Formula);
            Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Formula);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            string output_filename = TestGlobals.TestHelper.GetTestMethodOutputFilename(".vsd");
            doc.SaveAs(output_filename);
            doc.Close();
        }

        private DG.Drawing create_sample_graph()
        {
            var d = new DG.Drawing();

            var basic_stencil = "basic_u.vss";
            var n0 = d.AddShape("n0", "Node 0", basic_stencil, "Rectangle");
            n0.Size = new VA.Drawing.Size(3, 2);
            var n1 = d.AddShape("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = d.AddShape("n2", "Node 2", basic_stencil, "Rectangle");
            var n3 = d.AddShape("n3", "Node 3", basic_stencil, "Rectangle");
            var n4 = d.AddShape("n4", "Node 4\nUnconnected", basic_stencil, "Rectangle");

            var c0 = d.AddConnection("c0", n0, n1, "0 -> 1", VACONNECT.ConnectorType.Curved);
            var c1 = d.AddConnection("c1", n1, n2, "1 -> 2", VACONNECT.ConnectorType.RightAngle);
            var c2 = d.AddConnection("c2", n1, n0, "0 -> 1", VACONNECT.ConnectorType.Curved);
            var c3 = d.AddConnection("c3", n0, n2, "0 -> 2", VACONNECT.ConnectorType.Straight);
            var c4 = d.AddConnection("c4", n2, n3, "2 -> 3", VACONNECT.ConnectorType.Curved);
            var c5 = d.AddConnection("c5", n3, n0, "3 -> 0", VACONNECT.ConnectorType.Curved);

            return d;
        }
    }
}
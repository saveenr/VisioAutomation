using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes.Connections;
using VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using DG = VisioAutomation.Models.DirectedGraph;

namespace TestVisioAutomation
{
    [TestClass]
    public class DirectedGraph_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void DirectedGraph_WithBezierConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();
            
            var options = new DG.MSAGLLayoutOptions();
            options.UseDynamicConnectors = false;
            
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;
            directed_graph_drawing.Render(page,options);
            doc.Close(true);
        }

        [TestMethod]
        public void DirectedGraph_WithDynamicConnectors()
        {
            var directed_graph_drawing = this.create_sample_graph();

            var options = new VA.Models.DirectedGraph.MSAGLLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;
            
            directed_graph_drawing.Render(page1,options);
            doc.Close(true);
        }

        [TestMethod]
        public void RenderDirectedGraphWithCustomProps()
        {
            var d = new VA.Models.DirectedGraph.Drawing();

            var n0 = d.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            n0.CustomProperties = new Dictionary<string, CustomPropertyCells>();
            n0.CustomProperties["p1"] = new CustomPropertyCells("v1");
            n0.CustomProperties["p2"] = new CustomPropertyCells("v2");
            n0.CustomProperties["p3"] = new CustomPropertyCells("v3");

            var options = new VA.Models.DirectedGraph.MSAGLLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            d.Render(page1, options);
            
            Assert.IsNotNull(n0.VisioShape);
            var props_dic = CustomPropertyHelper.Get(n0.VisioShape);
            Assert.IsTrue(props_dic.Count>=3);
            Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Formula);
            Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Formula);
            Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Formula);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            doc.Close(true);
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

            var c0 = d.Connect("c0", n0, n1, "0 -> 1", ConnectorType.Curved);
            var c1 = d.Connect("c1", n1, n2, "1 -> 2", ConnectorType.RightAngle);
            var c2 = d.Connect("c2", n1, n0, "0 -> 1", ConnectorType.Curved);
            var c3 = d.Connect("c3", n0, n2, "0 -> 2", ConnectorType.Straight);
            var c4 = d.Connect("c4", n2, n3, "2 -> 3", ConnectorType.Curved);
            var c5 = d.Connect("c5", n3, n0, "3 -> 0", ConnectorType.Curved);

            return d;
        }
    }
}
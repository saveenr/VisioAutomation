using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Linq;
using DG = VisioAutomation.Layout.Models.DirectedGraph;

namespace TestVisioAutomation
{
    [TestClass]
    public class DirectedGraphLayoutTests : VisioAutomationTest
    {
        [TestMethod]
        public void RenderDirectedGraphWithBezierConnectors()
        {
            var directed_graph_drawing = new DG.Drawing();
            DG.Shape[] shapes;
            DG.Connector[] connectors;
            this.filldrawing(directed_graph_drawing,out shapes,out connectors);
            
            var options = new DG.MSAGLLayoutOptions();
            options.UseDynamicConnectors = false;
            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;

            directed_graph_drawing.Render(page,options);


            doc.Close(true);
        }

        private void filldrawing(DG.Drawing directed_graph_drawing, out DG.Shape[] nodes, out DG.Connector[] connectors)
        {
            var basic_stencil = "basic_u.vss";
            var n0 = directed_graph_drawing.AddShape("n0", "Node 0", basic_stencil, "Rectangle");
            n0.Size = new VA.Drawing.Size(3, 2);
            var n1 = directed_graph_drawing.AddShape("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = directed_graph_drawing.AddShape("n2", "Node 2", basic_stencil, "Rectangle");
            var n3 = directed_graph_drawing.AddShape("n3", "Node 3", basic_stencil, "Rectangle");
            var n4 = directed_graph_drawing.AddShape("n4", "Node 4\nUnconnected", basic_stencil, "Rectangle");

            nodes = new DG.Shape[]
            {
                n0,n1,n2,n3,n4
            };

            var c0 = directed_graph_drawing.Connect("c0", n0, n1, "0 -> 1", VA.Connections.ConnectorType.Curved);
            var c1 = directed_graph_drawing.Connect("c1", n1, n2, "1 -> 2", VA.Connections.ConnectorType.RightAngle);
            var c2 = directed_graph_drawing.Connect("c2", n1, n0, "0 -> 1", VA.Connections.ConnectorType.Curved);
            var c3 = directed_graph_drawing.Connect("c3", n0, n2, "0 -> 2", VA.Connections.ConnectorType.Straight);
            var c4 = directed_graph_drawing.Connect("c4", n2, n3, "2 -> 3", VA.Connections.ConnectorType.Curved);
            var c5 = directed_graph_drawing.Connect("c5", n3, n0, "3 -> 0", VA.Connections.ConnectorType.Curved);

            connectors = new DG.Connector[]
            {
                c0,c1,c2,c3,c4,c5
            };
        }

        [TestMethod]
        public void RenderDirectedGraphWithDynamicConnectors()
        {
            var directed_graph_drawing = new DG.Drawing();
            DG.Shape[] shapes;
            DG.Connector[] connectors;
            this.filldrawing(directed_graph_drawing, out shapes, out connectors);

            var options = new VA.Layout.Models.DirectedGraph.MSAGLLayoutOptions();
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
            var directed_graph_drawing = new VA.Layout.Models.DirectedGraph.Drawing();

            var n0 = directed_graph_drawing.AddShape("n0", "Untitled Node", "basflo_u.vss",
                                   "Decision");
            n0.Size = new VA.Drawing.Size(3, 2);
            n0.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
            n0.CustomProperties["p1"] = new VA.CustomProperties.CustomPropertyCells("v1");
            n0.CustomProperties["p2"] = new VA.CustomProperties.CustomPropertyCells("v2");
            n0.CustomProperties["p3"] = new VA.CustomProperties.CustomPropertyCells("v3");

            var options = new VA.Layout.Models.DirectedGraph.MSAGLLayoutOptions();
            options.UseDynamicConnectors = true;

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            directed_graph_drawing.Render(page1, options);
            
            Assert.IsNotNull(n0.VisioShape);
            var props_dic = VA.CustomProperties.CustomPropertyHelper.Get(n0.VisioShape);
            Assert.IsTrue(props_dic.Count>=3);
            Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Formula);
            Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Formula);
            Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Formula);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            doc.Close(true);
        }
    }
}
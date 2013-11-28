using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Shapes.Connections;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class PathAnalysis_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void PathAnalysis_VerifyEdgesAndTransitiveClosure()
        {
            var page1 = GetNewPage();
            var connectors_stencil = page1.Application.Documents.OpenStencil("connec_u.vss");
            var connectors_masters = connectors_stencil.Masters;

            var dcm = connectors_masters["Dynamic Connector"];

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "A";

            var s2 = page1.DrawRectangle(2, 0, 3, 1);
            s2.Text = "B";

            var c1 = page1.Drop(dcm, new VA.Drawing.Point(-2, -2));
            c1.Text = "Con";

            ConnectorHelper.ConnectShapes(s1, s2, c1);

            var edges1 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreBidirectional);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(2,map1.CountFromNodes());
            Assert.IsTrue(map1.HasConnectionFromTo("A","B"));
            Assert.IsTrue(map1.HasConnectionFromTo("B","A"));
            Assert.AreEqual(1,map1.CountConnectionsFrom("A"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("B"));

            var edges2 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            var map2 = new ConnectivityMap(edges2);
            Assert.AreEqual(0, map2.CountFromNodes());

            var edges3 = PathAnalysis.GetTransitiveClosure(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreBidirectional);
            var map3 = new ConnectivityMap(edges3);
            Assert.AreEqual(2, map3.CountFromNodes());
            Assert.IsTrue(map3.HasConnectionFromTo("A", "B"));
            Assert.IsTrue(map3.HasConnectionFromTo("B", "A"));
            Assert.AreEqual(1, map3.CountConnectionsFrom("A"));
            Assert.AreEqual(1, map3.CountConnectionsFrom("B"));

            var edges4 = PathAnalysis.GetTransitiveClosure(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            var map4  = new ConnectivityMap(edges4);
            Assert.AreEqual(0, map4.CountFromNodes());

            page1.Delete(0);
        }
    }
}
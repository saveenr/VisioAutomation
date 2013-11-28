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
        private IVisio.VisAutoConnectDir connect_dir_none = IVisio.VisAutoConnectDir.visAutoConnectDirNone;

        private void connect(IVisio.Shape a, IVisio.Shape b, bool a_arrow, bool b_arrow)
        {
            a.AutoConnect(b, connect_dir_none, null);
        }

        [TestMethod]
        public void Connects_EnumerableExtensionMethod()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            connect(shapes[0], shapes[1], false, false);
            connect(shapes[1], shapes[2], false, false);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);
        }

        [TestMethod]
        public void PathAnalysis_GetDirectEdgesRaw()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            connect(shapes[0], shapes[1], false, false);
            connect(shapes[1], shapes[2], false, false);

            var edges = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.Raw);
            var map = new ConnectivityMap(edges);
            Assert.AreEqual(2, map.CountFromNodes());
            Assert.IsTrue(map.HasConnectionFromTo("A","B"));
            Assert.IsTrue(map.HasConnectionFromTo("B", "C"));
            Assert.AreEqual(1, map.CountConnectionsFrom("A"));
            Assert.AreEqual(1, map.CountConnectionsFrom("B"));
            page1.Delete(0);
        }

        [TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreBidirectional()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            connect(shapes[0], shapes[1], false, false);
            connect(shapes[1], shapes[2], false, false);

            var edges1 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreBidirectional);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(3, map1.CountFromNodes());
            Assert.IsTrue(map1.HasConnectionFromTo("A", "B"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "C"));
            Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("A"));
            Assert.AreEqual(2, map1.CountConnectionsFrom("B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("C"));


            var edges2 = PathAnalysis.GetTransitiveClosure(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreBidirectional);
            var map2 = new ConnectivityMap(edges2);
            Assert.AreEqual(3, map2.CountFromNodes());
            Assert.IsTrue(map2.HasConnectionFromTo("A", "B"));
            Assert.IsTrue(map2.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map2.HasConnectionFromTo("B", "C"));
            Assert.IsTrue(map2.HasConnectionFromTo("C", "B"));
            Assert.IsTrue(map2.HasConnectionFromTo("A", "C"));
            Assert.IsTrue(map2.HasConnectionFromTo("C", "A"));
            
            Assert.AreEqual(2, map2.CountConnectionsFrom("A"));
            Assert.AreEqual(2, map2.CountConnectionsFrom("B"));
            Assert.AreEqual(2, map2.CountConnectionsFrom("C"));


            page1.Delete(0);
        }

        [TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            connect(shapes[0], shapes[1], false, false);
            connect(shapes[1], shapes[2], false, false);

            var edges1 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(0, map1.CountFromNodes());

            var edges2 = PathAnalysis.GetTransitiveClosure(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            var map2 = new ConnectivityMap(edges2);
            Assert.AreEqual(0, map2.CountFromNodes());

            page1.Delete(0);
        }

        private IVisio.Shape[] draw_standard_shapes(IVisio.Page page1)
        {
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(0, 3, 1, 4);
            var s3 = page1.DrawRectangle(3, 0, 4, 1);
            s1.Text = "A";
            s2.Text = "B";
            s3.Text = "C";
            return new[] { s1, s2, s3 };
        }
    }
}
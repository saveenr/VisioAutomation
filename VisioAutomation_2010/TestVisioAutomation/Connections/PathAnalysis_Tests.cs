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

        private IVisio.VisAutoConnectDir connect_dir_none = IVisio.VisAutoConnectDir.visAutoConnectDirNone;

        [TestMethod]
        public void Connects_EnumerableExtensionMethod()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            shapes[0].AutoConnect(shapes[1], connect_dir_none, null);
            shapes[1].AutoConnect(shapes[2], connect_dir_none, null);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);
        }

        [TestMethod]
        public void PathAnalysis_GetDirectEdgesRaw()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            shapes[0].AutoConnect(shapes[1], connect_dir_none, null); // RAW A->B
            shapes[1].AutoConnect(shapes[2], connect_dir_none, null); // RAW B->C

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

            shapes[0].AutoConnect(shapes[1], connect_dir_none, null);  // RAW A->B
            shapes[1].AutoConnect(shapes[2], connect_dir_none, null);  // RAW B->C

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
        public void Connects_GetDirectedEdgesExcludeNoArrows()
        {
            var page1 = GetNewPage();

            var shapes = draw_standard_shapes(page1);
            short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            var application = page1.Application;
            var documents = application.Documents;
            var connectors_stencil = documents.OpenEx("connec_u.vss", flags);
            var connectors_masters = connectors_stencil.Masters;

            var master = connectors_masters["Dynamic Connector"];

            var c1 = page1.Drop(master, -1, -1);
            connect(c1, shapes[0], shapes[1]);

            var c2 = page1.Drop(master, -1, -1);
            connect(c2, shapes[1], shapes[2]);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);

            var edges0 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            Assert.AreEqual(0, edges0.Count);

            var src_beginarrow = VA.ShapeSheet.SRCConstants.BeginArrow;
            var src_endarrow = VA.ShapeSheet.SRCConstants.EndArrow;

            var cell_beginarrow = c1.CellsSRC[src_beginarrow.Section, src_beginarrow.Row, src_beginarrow.Cell];
            var cell_endarow = c2.CellsSRC[src_endarrow.Section, src_endarrow.Row, src_endarrow.Cell];

            cell_beginarrow.FormulaU = "1";
            cell_endarow.FormulaU = "1";
            var edges1 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded);
            Assert.AreEqual(2, edges1.Count);
            Assert.AreEqual("B", edges1[0].From.Text);
            Assert.AreEqual("A", edges1[0].To.Text);
            Assert.AreEqual("B", edges1[1].From.Text);
            Assert.AreEqual("C", edges1[1].To.Text);

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

        private void connect(IVisio.Shape c1, IVisio.Shape from, IVisio.Shape to)
        {
            ConnectorHelper.ConnectShapes(from, to, c1);
        }

    }
}
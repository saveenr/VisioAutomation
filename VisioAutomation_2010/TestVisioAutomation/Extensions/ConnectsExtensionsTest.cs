using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VisioAutomation.Shapes.Connections;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ConnectsExtensionsTest : VisioAutomationTest
    {
        private IVisio.VisAutoConnectDir connect_dir_none = IVisio.VisAutoConnectDir.visAutoConnectDirNone;

        [TestMethod]
        public void Connects_GetUndirectedEdges()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            shapes[0].AutoConnect(shapes[1], connect_dir_none, null);
            shapes[1].AutoConnect(shapes[2], connect_dir_none, null);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);

            var edges = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.Raw);
            Assert.AreEqual(2, edges.Count);

            Assert.AreEqual(shapes[0], edges[0].From);
            Assert.AreEqual(shapes[1], edges[0].To);

            Assert.AreEqual(shapes[1], edges[1].From);
            Assert.AreEqual(shapes[2], edges[1].To);

            page1.Delete(0);
        }

        [TestMethod]
        public void Connects_GetDirectedEdgesIncludeNoArrowsAreBidirectional()
        {
            var page1 = GetNewPage();
            var shapes = draw_standard_shapes(page1);

            shapes[0].AutoConnect(shapes[1], connect_dir_none, null);
            shapes[1].AutoConnect(shapes[2], connect_dir_none, null);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);

            var edges = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.Arrows_NoArrowsAreBidirectional);
            Assert.AreEqual(4, edges.Count);

            Assert.AreEqual(shapes[1], edges[0].From);
            Assert.AreEqual(shapes[0], edges[0].To);

            Assert.AreEqual(shapes[0], edges[1].From);
            Assert.AreEqual(shapes[1], edges[1].To);

            Assert.AreEqual(shapes[2], edges[2].From);
            Assert.AreEqual(shapes[1], edges[2].To);

            Assert.AreEqual(shapes[1], edges[3].From);
            Assert.AreEqual(shapes[2], edges[3].To);

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
            var connectors_stencil = documents.OpenEx("connec_u.vss",flags);
            var connectors_masters = connectors_stencil.Masters;

            var master = connectors_masters["Dynamic Connector"];

            var c1 = page1.Drop(master, -1, -1);
            connect(c1, shapes[0], shapes[1]);
            
            var c2 = page1.Drop(master, -1, -1);
            connect(c2, shapes[1], shapes[2]);

            var cons = page1.Connects.AsEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);

            var edges0 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.Arrows_NoArrowsAreExcluded);
            Assert.AreEqual(0, edges0.Count);

            var src_beginarrow = VA.ShapeSheet.SRCConstants.BeginArrow;
            var src_endarrow = VA.ShapeSheet.SRCConstants.EndArrow;

            var cell_beginarrow = c1.CellsSRC[src_beginarrow.Section, src_beginarrow.Row, src_beginarrow.Cell];
            var cell_endarow = c2.CellsSRC[src_endarrow.Section, src_endarrow.Row, src_endarrow.Cell];

            cell_beginarrow.FormulaU = "1";
            cell_endarow.FormulaU = "1";
            var edges1 = PathAnalysis.GetDirectedEdges(page1, DirectedEdgeHandling.Arrows_NoArrowsAreExcluded);
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
            return new[] {s1, s2, s3};
        }

        private void connect(IVisio.Shape c1, IVisio.Shape from, IVisio.Shape to)
        {
            ConnectorHelper.ConnectShapes(from, to, c1);
        }
    }
}
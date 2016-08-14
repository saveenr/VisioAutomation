using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes.Connectors;
using VisioAutomation.ShapeSheet.Update;
using VADRAW=VisioAutomation.Drawing;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Tests.DocumentAnalysis
{
    [TestClass]
    public class PathAnalysis_Tests : VisioAutomationTest
    {
        private void connect(IVisio.Shape a, IVisio.Shape b, bool a_arrow, bool b_arrow)
        {
            var page = a.ContainingPage;
            var connectors_stencil = page.Application.Documents.OpenStencil("connec_u.vss");
            var connectors_masters = connectors_stencil.Masters;

            var dcm = connectors_masters["Dynamic Connector"];

            var drop_point = new VADRAW.Point(-2, -2);
            var c1 = page.Drop(dcm, drop_point);
            ConnectorHelper.ConnectShapes(a, b, c1);

            //a.AutoConnect(b, connect_dir_none, null);

            if (a_arrow || b_arrow)
            {
                var update = new UpdateSIDSRCFormula();
                if (a_arrow)
                {
                    update.SetFormula(c1.ID16, VASS.SRCConstants.BeginArrow, "13");                    
                }
                if (b_arrow)
                {
                    update.SetFormula(c1.ID16, VASS.SRCConstants.EndArrow, "13");
                }
                update.Execute(page);
            }
        }

        [TestMethod]
        public void Connects_EnumerableExtensionMethod()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var cons = page1.Connects.ToEnumerable().ToList();
            Assert.AreEqual(4, cons.Count);
            page1.Delete(0);
        }

        [TestMethod]
        public void PathAnalysis_GetDirectEdgesRaw()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var ch = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch.DirectionSource = VisioAutomation.DocumentAnalysis.DirectionSource.UseConnectionOrder;

            var edges = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(page1, ch);
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
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var ch1 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch1.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.TreatAsBidirectional;
            var edges1 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(page1, ch1);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(3, map1.CountFromNodes());
            Assert.IsTrue(map1.HasConnectionFromTo("A", "B"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "C"));
            Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("A"));
            Assert.AreEqual(2, map1.CountConnectionsFrom("B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var ch2 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch2.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.TreatAsBidirectional;

            var edges2 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(page1, ch2);
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
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var ch = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.Exclude;

            var edges1 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(page1, ch);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(0, map1.CountFromNodes());

            var edges2 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(page1, ch);
            var map2 = new ConnectivityMap(edges2);
            Assert.AreEqual(0, map2.CountFromNodes());

            page1.Delete(0);
        }

        [TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded_withArrowHeads1()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], true, false);
            this.connect(shapes[1], shapes[2], true, false);

            var ch1 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch1.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.Exclude;

            var edges1 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(page1, ch1);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(2, map1.CountFromNodes());
            Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var ch2 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch2.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.TreatAsBidirectional;


            var edges2 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(page1, ch2);
            var map2 = new ConnectivityMap(edges1);
            Assert.AreEqual(2, map2.CountFromNodes());
            Assert.IsTrue(map2.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map2.HasConnectionFromTo("C", "B"));
            Assert.AreEqual(1, map2.CountConnectionsFrom("B"));
            Assert.AreEqual(1, map2.CountConnectionsFrom("C"));


            page1.Delete(0);
        }

        [TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded_withArrowHeads2()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], true, true);
            this.connect(shapes[1], shapes[2], true, true);

            var ch1 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch1.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.Exclude;

            var edges1 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetDirectedEdges(page1, ch1);
            var map1 = new ConnectivityMap(edges1);
            Assert.AreEqual(3, map1.CountFromNodes());
            Assert.IsTrue(map1.HasConnectionFromTo("A", "B"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            Assert.IsTrue(map1.HasConnectionFromTo("B", "C"));
            Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("A"));
            Assert.AreEqual(2, map1.CountConnectionsFrom("B"));
            Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var ch2 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            ch2.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.TreatAsBidirectional;

            var edges2 = VisioAutomation.DocumentAnalysis.ConnectionAnalyzer.GetTransitiveClosure(page1, ch2);
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
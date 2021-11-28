using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using VASS=VisioAutomation.ShapeSheet;
using VADRAW = VisioAutomation.Core;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Analyzers
{
    [MUT.TestClass]
    public class ConnectionTests : Framework.VTest
    {
        private void connect(IVisio.Shape a, IVisio.Shape b, bool a_arrow, bool b_arrow)
        {
            var page = a.ContainingPage;
            var connectors_stencil = page.Application.Documents.OpenStencil("connec_u.vss");
            var connectors_masters = connectors_stencil.Masters;

            var dcm = connectors_masters["Dynamic Connector"];

            var drop_point = new VADRAW.Point(-2, -2);
            var c1 = page.Drop(dcm, drop_point);
            VA.Shapes.ConnectorHelper.ConnectShapes(a, b, c1);

            //a.AutoConnect(b, connect_dir_none, null);

            if (a_arrow || b_arrow)
            {
                var writer = new VASS.Writers.SidSrcWriter();
                if (a_arrow)
                {
                    writer.SetValue(c1.ID16, VisioAutomation.Core.SrcConstants.LineBeginArrow, "13");                    
                }
                if (b_arrow)
                {
                    writer.SetValue(c1.ID16, VisioAutomation.Core.SrcConstants.LineEndArrow, "13");
                }

                writer.Commit(page, VisioAutomation.Core.CellValueType.Formula);
            }
        }

        [MUT.TestMethod]
        public void Connects_EnumerableExtensionMethod()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var cons = page1.Connects.ToList();
            MUT.Assert.AreEqual(4, cons.Count);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void PathAnalysis_GetDirectEdgesRaw()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var options = new VA.Analyzers.ConnectionAnalyzerOptions();
            options.EdgeDirectionSource = VA.Analyzers.EdgeDirectionSource.UseConnectionOrder;

            var edges = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page1, options);
            var map = new ConnectivityMap(edges);
            MUT.Assert.AreEqual(2, map.CountFromNodes());
            MUT.Assert.IsTrue(map.HasConnectionFromTo("A","B"));
            MUT.Assert.IsTrue(map.HasConnectionFromTo("B", "C"));
            MUT.Assert.AreEqual(1, map.CountConnectionsFrom("A"));
            MUT.Assert.AreEqual(1, map.CountConnectionsFrom("B"));
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreBidirectional()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var options1 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options1.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.IncludeEdgesForBothDirections;
            var edges1 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page1, options1);
            var map1 = new ConnectivityMap(edges1);
            MUT.Assert.AreEqual(3, map1.CountFromNodes());
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("A", "B"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("B", "C"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("A"));
            MUT.Assert.AreEqual(2, map1.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var options2 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options2.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.IncludeEdgesForBothDirections;

            var edges2 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdgesTransitive(page1, options2);
            var map2 = new ConnectivityMap(edges2);
            MUT.Assert.AreEqual(3, map2.CountFromNodes());
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("A", "B"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("B", "C"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("C", "B"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("A", "C"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("C", "A"));
            
            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("A"));
            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("C"));


            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], false, false);
            this.connect(shapes[1], shapes[2], false, false);

            var options = new VA.Analyzers.ConnectionAnalyzerOptions();
            options.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.ExcludeEdge;

            var edges1 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page1, options);
            var map1 = new ConnectivityMap(edges1);
            MUT.Assert.AreEqual(0, map1.CountFromNodes());

            var edges2 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdgesTransitive(page1, options);
            var map2 = new ConnectivityMap(edges2);
            MUT.Assert.AreEqual(0, map2.CountFromNodes());

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded_withArrowHeads1()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], true, false);
            this.connect(shapes[1], shapes[2], true, false);

            var options1 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options1.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.ExcludeEdge;

            var edges1 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page1, options1);
            var map1 = new ConnectivityMap(edges1);
            MUT.Assert.AreEqual(2, map1.CountFromNodes());
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var options2 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options2.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.IncludeEdgesForBothDirections;


            var edges2 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdgesTransitive(page1, options2);
            var map2 = new ConnectivityMap(edges1);
            MUT.Assert.AreEqual(2, map2.CountFromNodes());
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("C", "B"));
            MUT.Assert.AreEqual(1, map2.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(1, map2.CountConnectionsFrom("C"));


            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Connects_GetDirectedEdges_EdgesWithoutArrowsAreExcluded_withArrowHeads2()
        {
            var page1 = this.GetNewPage();
            var shapes = this.draw_standard_shapes(page1);

            this.connect(shapes[0], shapes[1], true, true);
            this.connect(shapes[1], shapes[2], true, true);

            var options1 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options1.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.ExcludeEdge;

            var edges1 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdges(page1, options1);
            var map1 = new ConnectivityMap(edges1);
            MUT.Assert.AreEqual(3, map1.CountFromNodes());
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("A", "B"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("B", "C"));
            MUT.Assert.IsTrue(map1.HasConnectionFromTo("C", "B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("A"));
            MUT.Assert.AreEqual(2, map1.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(1, map1.CountConnectionsFrom("C"));

            var options2 = new VA.Analyzers.ConnectionAnalyzerOptions();
            options2.EdgeNoArrowsHandling = VA.Analyzers.EdgeNoArrowsHandling.IncludeEdgesForBothDirections;

            var edges2 = VA.Analyzers.ConnectionAnalyzer.GetDirectedEdgesTransitive(page1, options2);
            var map2 = new ConnectivityMap(edges2);
            MUT.Assert.AreEqual(3, map2.CountFromNodes());
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("A", "B"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("B", "A"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("B", "C"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("C", "B"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("A", "C"));
            MUT.Assert.IsTrue(map2.HasConnectionFromTo("C", "A"));

            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("A"));
            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("B"));
            MUT.Assert.AreEqual(2, map2.CountConnectionsFrom("C"));


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
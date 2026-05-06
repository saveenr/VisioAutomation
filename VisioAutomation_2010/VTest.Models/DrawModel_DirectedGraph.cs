using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VTest.Models
{
    [MUT.TestClass]
    public class DrawModel_DirectedGraph : Framework.VTest
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
            
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(DirectedGraph_WithBezierConnectors),".vsd");
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

            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(DirectedGraph_WithDynamicConnectors),".vsd");
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
            n0.CustomProperties = new VisioAutomation.Shapes.CustomPropertyDictionary();
            n0.CustomProperties["p1"] = new VisioAutomation.Shapes.CustomPropertyCells("\"v1\"");
            n0.CustomProperties["p2"] = new VisioAutomation.Shapes.CustomPropertyCells("\"v2\"");
            n0.CustomProperties["p3"] = new VisioAutomation.Shapes.CustomPropertyCells("\"v3\"");

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page1 = visapp.ActivePage;

            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = true;
            renderer.Render(page1, d);

            MUT.Assert.IsNotNull(n0.VisioShape);
            var props_dic = VisioAutomation.Shapes.CustomPropertyHelper.GetDictionary(n0.VisioShape, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.IsTrue(props_dic.Count>=3);
            MUT.Assert.AreEqual("\"v1\"",props_dic["p1"].Value.Value);
            MUT.Assert.AreEqual("\"v2\"", props_dic["p2"].Value.Value);
            MUT.Assert.AreEqual("\"v3\"", props_dic["p3"].Value.Value);

            page1.Application.ActiveWindow.ViewFit = (short) IVisio.VisWindowFit.visFitPage;

            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(RenderDirectedGraphWithCustomProps),".vsd");
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

        [MUT.TestMethod]
        public void Scripting_Draw_DirectedGraph1()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_1.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph1), ".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Draw_DirectedGraph2()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_2.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph2), ".vsd");
            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Draw_DirectedGraph3()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_3.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph3), ".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Draw_DirectedGraph4()
        {
            // Load the graph
            string xml = this.get_datafile_content(@"datafiles\directed_graph_4.xml");

            // Draw the graph
            var client = this.GetScriptingClient();
            this.draw_directed_graph(client, xml);

            // Cleanup
            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Draw_DirectedGraph4), ".vsd");

            client.Document.SaveDocumentAs(VisioScripting.TargetDocument.Auto, output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }


        [MUT.TestMethod]
        public void Loader_ConnectorType_DefaultsToCurvedWhenAttributeMissing()
        {
            var dg = this.load_two_node_graph(connectortype_attr: null);
            MUT.Assert.AreEqual(VA.Models.ConnectorType.Curved, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorType_StraightFromXml()
        {
            var dg = this.load_two_node_graph(connectortype_attr: "Straight");
            MUT.Assert.AreEqual(VA.Models.ConnectorType.Straight, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorType_RightAngleFromXml()
        {
            var dg = this.load_two_node_graph(connectortype_attr: "RightAngle");
            MUT.Assert.AreEqual(VA.Models.ConnectorType.RightAngle, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorType_CurvedFromXml()
        {
            var dg = this.load_two_node_graph(connectortype_attr: "Curved");
            MUT.Assert.AreEqual(VA.Models.ConnectorType.Curved, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorType_UnrecognizedValueThrows()
        {
            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => this.load_two_node_graph(connectortype_attr: "Wiggly"));
        }

        [MUT.TestMethod]
        public void Loader_Direction_DefaultsToTopToBottomWhenAttributeMissing()
        {
            var dg = this.load_two_node_graph_with_renderoptions("");
            MUT.Assert.AreEqual(VADG.MsaglDirection.TopToBottom, dg.Layouts[0].LayoutOptions.Direction);
        }

        [MUT.TestMethod]
        public void Loader_Direction_LeftToRightFromXml()
        {
            var dg = this.load_two_node_graph_with_renderoptions(" direction=\"LeftToRight\"");
            MUT.Assert.AreEqual(VADG.MsaglDirection.LeftToRight, dg.Layouts[0].LayoutOptions.Direction);
        }

        [MUT.TestMethod]
        public void Loader_Direction_RightToLeftFromXml()
        {
            var dg = this.load_two_node_graph_with_renderoptions(" direction=\"RightToLeft\"");
            MUT.Assert.AreEqual(VADG.MsaglDirection.RightToLeft, dg.Layouts[0].LayoutOptions.Direction);
        }

        [MUT.TestMethod]
        public void Loader_Direction_BottomToTopFromXml()
        {
            var dg = this.load_two_node_graph_with_renderoptions(" direction=\"BottomToTop\"");
            MUT.Assert.AreEqual(VADG.MsaglDirection.BottomToTop, dg.Layouts[0].LayoutOptions.Direction);
        }

        [MUT.TestMethod]
        public void Loader_Direction_UnrecognizedValueThrows()
        {
            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => this.load_two_node_graph_with_renderoptions(" direction=\"Diagonal\""));
        }

        [MUT.TestMethod]
        public void Loader_Layout_SugiyamaIsAccepted()
        {
            var dg = this.load_two_node_graph_with_renderoptions(" layout=\"Sugiyama\"");
            MUT.Assert.AreEqual(1, dg.Layouts.Count);
        }

        [MUT.TestMethod]
        public void Loader_Layout_UnrecognizedValueThrows()
        {
            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => this.load_two_node_graph_with_renderoptions(" layout=\"Foo\""));
        }

        [MUT.TestMethod]
        public void Loader_RootElement_WrongNameThrows()
        {
            string xml = "<wrongroot><page>" +
                "<renderoptions usedynamicconnectors=\"true\" scalingfactor=\"20\" />" +
                "<shapes><shape id=\"n1\" label=\"A\" stencil=\"basic_u.vss\" master=\"Rectangle\" /></shapes>" +
                "<connectors></connectors>" +
                "</page></wrongroot>";
            var dg_xml = SXL.XDocument.Parse(xml);
            var client = this.GetScriptingClient();
            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => VisioScripting.Loaders.DirectedGraphDocumentLoader.LoadFromXml(client, dg_xml));
        }

        [MUT.TestMethod]
        public void DirectedGraph_LeftToRight_RendersHorizontally()
        {
            var dg = new VADG.DirectedGraphLayout();
            dg.LayoutOptions.Direction = VADG.MsaglDirection.LeftToRight;
            var basic_stencil = "basic_u.vss";
            var n0 = dg.AddNode("n0", "Node 0", basic_stencil, "Rectangle");
            var n1 = dg.AddNode("n1", "Node 1", basic_stencil, "Rectangle");
            var n2 = dg.AddNode("n2", "Node 2", basic_stencil, "Rectangle");
            dg.AddEdge("c0", n0, n1, "0 -> 1", VA.Models.ConnectorType.Curved);
            dg.AddEdge("c1", n1, n2, "1 -> 2", VA.Models.ConnectorType.Curved);

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;

            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions = dg.LayoutOptions;
            renderer.Render(page, dg);

            double pinx_n0 = n0.VisioShape.Cells["PinX"].ResultIU;
            double pinx_n1 = n1.VisioShape.Cells["PinX"].ResultIU;
            double pinx_n2 = n2.VisioShape.Cells["PinX"].ResultIU;
            double piny_n0 = n0.VisioShape.Cells["PinY"].ResultIU;
            double piny_n2 = n2.VisioShape.Cells["PinY"].ResultIU;

            // For LeftToRight, the chain should flow horizontally: each downstream node strictly to the right.
            MUT.Assert.IsTrue(pinx_n0 < pinx_n1, string.Format("Expected n0.PinX < n1.PinX, got {0} vs {1}", pinx_n0, pinx_n1));
            MUT.Assert.IsTrue(pinx_n1 < pinx_n2, string.Format("Expected n1.PinX < n2.PinX, got {0} vs {1}", pinx_n1, pinx_n2));
            // And the spread along Y should be small relative to the X spread.
            double dx = System.Math.Abs(pinx_n2 - pinx_n0);
            double dy = System.Math.Abs(piny_n2 - piny_n0);
            MUT.Assert.IsTrue(dx > dy, string.Format("Expected horizontal spread > vertical spread, got dx={0} dy={1}", dx, dy));

            doc.Close();
        }

        private VA.Models.Layouts.DirectedGraph.DirectedGraphDocument load_two_node_graph(string connectortype_attr)
        {
            string ct_attr = connectortype_attr == null ? "" : string.Format(" connectortype=\"{0}\"", connectortype_attr);
            return this.load_two_node_graph_with_renderoptions(ct_attr);
        }

        private VA.Models.Layouts.DirectedGraph.DirectedGraphDocument load_two_node_graph_with_renderoptions(string extra_attrs)
        {
            string xml = string.Format(
                "<directedgraph>" +
                "<page>" +
                "<renderoptions usedynamicconnectors=\"true\" scalingfactor=\"20\"{0} />" +
                "<shapes>" +
                "<shape id=\"n1\" label=\"A\" stencil=\"basic_u.vss\" master=\"Rectangle\" />" +
                "<shape id=\"n2\" label=\"B\" stencil=\"basic_u.vss\" master=\"Rectangle\" />" +
                "</shapes>" +
                "<connectors>" +
                "<connector id=\"c1\" from=\"n1\" to=\"n2\" label=\"\" />" +
                "</connectors>" +
                "</page>" +
                "</directedgraph>",
                extra_attrs ?? "");
            var dg_xml = SXL.XDocument.Parse(xml);
            var client = this.GetScriptingClient();
            return VisioScripting.Loaders.DirectedGraphDocumentLoader.LoadFromXml(client, dg_xml);
        }

        private void draw_directed_graph(VisioScripting.Client client, string dg_text)
        {
            var dg_xml = SXL.XDocument.Parse(dg_text);
            var dgdoc = VisioScripting.Loaders.DirectedGraphDocumentLoader.LoadFromXml(client, dg_xml);

            // TODO: Investigate if this this special case for Visio 2013 can be removed
            // this is a temporary fix to handle the fact that server_u.vss in Visio 2013 doesn't result in server_u.vssx 
            // getting automatically loaded

            var version = client.Application.ApplicationVersion;
            if (version.Major >= 15)
            {
                foreach (var drawing in dgdoc.Layouts)
                {
                    foreach (var shape in drawing.Nodes)
                    {
                        if (shape.StencilName == "server_u.vss")
                        {
                            shape.StencilName = "server_u.vssx";
                        }
                    }
                }
            }

            var dgstyling = new VA.Models.Layouts.DirectedGraph.DirectedGraphStyling();

            client.Model.DrawDirectedGraphDocument(dgdoc, dgstyling);
        }

    }
}
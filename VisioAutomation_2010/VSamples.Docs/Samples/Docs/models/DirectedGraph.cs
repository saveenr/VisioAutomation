using IVisio = Microsoft.Office.Interop.Visio;
using SXL = System.Xml.Linq;
using VA = VisioAutomation;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/directed-graph.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsDirectedGraphSamples
    {
        public static void Sample_01(IVisio.Page visioPage)
        {
            // <sample id="models/directed-graph#01" doc="models/directed-graph.md" line="23">
            var d = new VADG.DirectedGraphLayout();

            const string basic = "basic_u.vss";
            var n0 = d.AddNode("n0", "Node 0", basic, "Rectangle");
            n0.Size = new VA.Core.Size(3, 2);

            var n1 = d.AddNode("n1", "Node 1", basic, "Rectangle");
            var n2 = d.AddNode("n2", "Node 2", basic, "Rectangle");
            var n3 = d.AddNode("n3", "Node 3", basic, "Rectangle");
            d.AddNode("n4", "Node 4 (unconnected)", basic, "Rectangle");

            d.AddEdge("c0", n0, n1, "0 -> 1", VA.Models.ConnectorType.Curved);
            d.AddEdge("c1", n1, n2, "1 -> 2", VA.Models.ConnectorType.RightAngle);
            d.AddEdge("c2", n1, n0, "0 -> 1", VA.Models.ConnectorType.Curved);     // back-edge: cycle
            d.AddEdge("c3", n0, n2, "0 -> 2", VA.Models.ConnectorType.Straight);
            d.AddEdge("c4", n2, n3, "2 -> 3", VA.Models.ConnectorType.Curved);
            d.AddEdge("c5", n3, n0, "3 -> 0", VA.Models.ConnectorType.Curved);     // closes the cycle

            var renderer = new VADG.MsaglRenderer();
            renderer.Render(visioPage, d);
            // </sample>
        }

        public static void Sample_02(IVisio.Page visioPage, VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#02" doc="models/directed-graph.md" line="62">
            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = false;   // honor MSAGL's geometry
            renderer.Render(visioPage, d);
            // </sample>
        }

        public static void Sample_03(VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#03" doc="models/directed-graph.md" line="72">
            d.LayoutOptions.Direction = VADG.MsaglDirection.LeftToRight;
            // </sample>
        }

        public static void Sample_04(IVisio.Page visioPage, VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#04" doc="models/directed-graph.md" line="80">
            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions = d.LayoutOptions;
            renderer.Render(visioPage, d);
            // </sample>
        }

        public static void Sample_05(IVisio.Page visioPage, VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#05" doc="models/directed-graph.md" line="118">
            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions.Direction = VADG.MsaglDirection.LeftToRight;
            renderer.LayoutOptions.EdgeLabelBoxSize = new VA.Core.Size(0.8, 0.12);   // room for short labels only
            renderer.LayoutOptions.LayerSeparation = 0.25;                           // inches between layers
            renderer.Render(visioPage, d);
            // </sample>
        }

        public static void Sample_06(VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#06" doc="models/directed-graph.md" line="132">
            var n0 = d.AddNode("n0", "Untitled Node", "basflo_u.vss", "Decision");
            n0.Size = new VA.Core.Size(3, 2);
            n0.CustomProperties = new VisioAutomation.Shapes.CustomPropertyDictionary();
            n0.CustomProperties["p1"] = new VisioAutomation.Shapes.CustomPropertyCells("\"v1\"");
            n0.CustomProperties["p2"] = new VisioAutomation.Shapes.CustomPropertyCells("\"v2\"");
            // </sample>
        }

        public static void Sample_07(VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#07" doc="models/directed-graph.md" line="146">
            var n = d.AddNode("n", "Sized + styled", "basic_u.vss", "Rectangle");
            n.Size = new VA.Core.Size(1.25, 0.25);

            n.Cells = new VA.Models.Dom.ShapeCells();
            n.Cells.FillForeground = "rgb(0,255,0)";
            n.Cells.FillPattern = 1;
            // </sample>
        }

        public static void Sample_08(VisioScripting.Client client, string xml_string)
        {
            // <sample id="models/directed-graph#08" doc="models/directed-graph.md" line="163">
            var xml = SXL.XDocument.Parse(xml_string);
            var dg_doc = client.Model.LoadDirectedGraphFromXml(xml);

            var styling = new VA.Models.Layouts.DirectedGraph.DirectedGraphStyling();
            client.Model.DrawDirectedGraphDocument(dg_doc, styling);
            // </sample>
        }

        public static void Sample_09(IVisio.Page visioPage, VADG.DirectedGraphLayout d)
        {
            // <sample id="models/directed-graph#09" doc="models/directed-graph.md" line="185">
            var renderer = new VADG.VisioLayoutRenderer();
            renderer.LayoutOptions.VisioLayoutStyle = new VisioAutomation.Models.LayoutStyles.HierarchyLayoutStyle();
            renderer.Render(visioPage, d);
            // </sample>
        }
    }
}

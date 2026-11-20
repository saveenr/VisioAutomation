using System.Collections.Generic;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SXL = System.Xml.Linq;

namespace VTest.Models
{
    // The margin around a directed graph's page comes from two settings: a layout's LayoutOptions.PageBorderWidth
    // (per page) and the document's BorderSize (the default for every page). A layout that has set PageBorderWidth uses it;
    // otherwise the document's BorderSize applies.
    //
    // The tests compare margins with each other instead of with fixed numbers, because Visio adds a margin of its own
    // when it fits a page to its contents, and that does not depend on either setting.
    [MUT.TestClass]
    public class DirectedGraphPageBorderTests : Framework.VTest
    {
        private const double Tolerance = 0.01;

        [MUT.TestMethod]
        public void PageBorderWidth_IsNotSetUntilAssigned_AndStillReadsAsHalfAnInch()
        {
            var options = new VADG.MsaglOptions();
            MUT.Assert.IsFalse(options.IsPageBorderWidthSet);
            MUT.Assert.AreEqual(0.5, options.PageBorderWidth.Width);
            MUT.Assert.AreEqual(0.5, options.PageBorderWidth.Height);

            options.PageBorderWidth = new VA.Core.Size(0.25, 0.75);
            MUT.Assert.IsTrue(options.IsPageBorderWidthSet);
            MUT.Assert.AreEqual(0.25, options.PageBorderWidth.Width);
            MUT.Assert.AreEqual(0.75, options.PageBorderWidth.Height);

            MUT.Assert.IsFalse(new VADG.MsaglOptions().IsPageBorderWidthSet, "another instance is unaffected");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_LayoutWithoutPageBorderWidth_UsesTheDocumentBorderSize()
        {
            double none = this.draw_one_page(null, 0.0);
            double one = this.draw_one_page(null, 1.0);
            double two = this.draw_one_page(null, 2.0);

            MUT.Assert.AreEqual(1.0, one - none, Tolerance, "a document border of 1.0 should add 1.0 on each side");
            MUT.Assert.AreEqual(2.0, two - none, Tolerance, "a document border of 2.0 should add 2.0 on each side");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_DefaultSettings_GiveTheSameMarginAsAnExplicitOneInchBorderSize()
        {
            var layout = make_layout();
            var dgdoc = new VADG.DirectedGraphDocument();
            dgdoc.Layouts.Add(layout);
            double defaults = this.draw_and_measure(dgdoc)[0];

            double none = this.draw_one_page(null, 0.0);
            double explicit_one = this.draw_one_page(null, 1.0);

            MUT.Assert.AreEqual(explicit_one, defaults, Tolerance, "the default document border is 1.0");
            MUT.Assert.AreEqual(1.0, defaults - none, Tolerance, "an untouched layout should end up with the document's default 1.0 border, as before");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_LayoutPageBorderWidth_WinsOverTheDocumentBorderSize()
        {
            double none = this.draw_one_page(null, 0.0);
            double small = this.draw_one_page(0.25, 2.0);
            double large = this.draw_one_page(1.0, 2.0);

            MUT.Assert.AreEqual(0.25, small - none, Tolerance, "the layout's 0.25 should be used, not the document's 2.0");
            MUT.Assert.AreEqual(1.0, large - none, Tolerance, "the layout's 1.0 should be used, not the document's 2.0");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_LayoutPageBorderWidth_DoesNotDependOnTheDocumentBorderSize()
        {
            double with_small_document_border = this.draw_one_page(0.25, 0.0);
            double with_large_document_border = this.draw_one_page(0.25, 5.0);

            MUT.Assert.AreEqual(with_small_document_border, with_large_document_border, Tolerance);
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_PageBorderWidthSetToTheDefaultValue_StillCountsAsSet()
        {
            // Setting 0.5 (the value an unset layout reads as) is still an explicit choice, so it beats a document border of 2.0.
            double none = this.draw_one_page(null, 0.0);
            double half_set_explicitly = this.draw_one_page(0.5, 2.0);

            MUT.Assert.AreEqual(0.5, half_set_explicitly - none, Tolerance);
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_PageBorderWidthSetToZero_WinsOverTheDocumentBorderSize()
        {
            double none = this.draw_one_page(null, 0.0);
            double zero_set_explicitly = this.draw_one_page(0.0, 2.0);

            MUT.Assert.AreEqual(0.0, zero_set_explicitly - none, Tolerance, "an explicit zero border should remove the document's border for that page");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_TwoLayouts_EachPageFollowsItsOwnSetting()
        {
            double none = this.draw_one_page(null, 0.0);

            var first = make_layout();
            first.LayoutOptions.PageBorderWidth = new VA.Core.Size(0.25, 0.25);
            var second = make_layout();    // does not set its own border
            var dgdoc = new VADG.DirectedGraphDocument();
            dgdoc.Layouts.Add(first);
            dgdoc.Layouts.Add(second);
            dgdoc.BorderSize = new VA.Core.Size(2.0, 2.0);

            var margins = this.draw_and_measure(dgdoc);

            MUT.Assert.AreEqual(2, margins.Count, "one page per layout");
            MUT.Assert.AreEqual(0.25, margins[0] - none, Tolerance, "the first page should use its layout's border");
            MUT.Assert.AreEqual(2.0, margins[1] - none, Tolerance, "the second page should use the document's border");
        }

        [MUT.TestMethod]
        public void MsaglRenderer_UsedDirectly_StillUsesPageBorderWidth_AndAnUnsetOneReadsAsHalfAnInch()
        {
            double unset = this.render_directly(null);
            double explicit_half = this.render_directly(0.5);
            double quarter = this.render_directly(0.25);
            double one = this.render_directly(1.0);

            MUT.Assert.AreEqual(explicit_half, unset, Tolerance, "an unset border should behave as 0.5 when the renderer is used directly");
            MUT.Assert.AreEqual(0.75, one - quarter, Tolerance, "the renderer's own border should still decide the margin");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_PageBorderFromXml_WinsOverTheDocumentBorderFromXml()
        {
            double none = this.draw_one_page(null, 0.0);
            var dgdoc = load_xml(
                "<documentoptions borderwidth='2' borderheight='2' />",
                page_xml("borderwidth='0.25' borderheight='0.25'"));

            var margins = this.draw_and_measure(dgdoc);

            MUT.Assert.AreEqual(0.25, margins[0] - none, Tolerance, "the page's own border from the XML should be used, not the document's 2.0");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_TwoPagesFromXml_EachFollowsItsOwnBorder()
        {
            double none = this.draw_one_page(null, 0.0);
            var dgdoc = load_xml(
                "<documentoptions borderwidth='2' borderheight='2' />",
                page_xml("borderwidth='0.25' borderheight='0.25'"),
                page_xml(""));   // no border of its own, so it follows the document

            var margins = this.draw_and_measure(dgdoc);

            MUT.Assert.AreEqual(2, margins.Count);
            MUT.Assert.AreEqual(0.25, margins[0] - none, Tolerance, "the first page has its own border");
            MUT.Assert.AreEqual(2.0, margins[1] - none, Tolerance, "the second page follows the document's border");
        }

        [MUT.TestMethod]
        public void DrawDirectedGraphDocument_OnlyOnePageBorderSideInXml_TheOtherSideFollowsTheDocument()
        {
            // Measure the horizontal margin: width comes from the page, height from the document.
            double none = this.draw_one_page(null, 0.0);
            var dgdoc = load_xml(
                "<documentoptions borderwidth='2' borderheight='2' />",
                page_xml("borderwidth='0.25'"));

            var margins = this.draw_and_measure(dgdoc);

            MUT.Assert.AreEqual(0.25, margins[0] - none, Tolerance, "the width the page gives should be used");
        }

        // -- helpers ---------------------------------------------------------------------------------

        private static string page_xml(string render_attrs)
        {
            return
                "<page>" +
                "<renderoptions usedynamicconnectors='true' scalingfactor='20' " + render_attrs + " />" +
                "<shapes>" +
                "<shape id='n1' label='A' stencil='basic_u.vss' master='Rectangle' />" +
                "<shape id='n2' label='B' stencil='basic_u.vss' master='Rectangle' />" +
                "</shapes>" +
                "<connectors><connector id='c1' from='n1' to='n2' label='' /></connectors>" +
                "</page>";
        }

        private VADG.DirectedGraphDocument load_xml(string root_children, params string[] pages)
        {
            string xml = "<directedgraph>" + root_children + string.Join("", pages) + "</directedgraph>";
            return this.GetScriptingClient().Model.LoadDirectedGraphFromXml(SXL.XDocument.Parse(xml));
        }


        private static VADG.DirectedGraphLayout make_layout()
        {
            var layout = new VADG.DirectedGraphLayout();
            var nodes = new List<VADG.Node>();
            foreach (var i in new[] { 0, 1, 2, 3 })
            {
                nodes.Add(layout.AddNode("n" + i, "Node " + i, "basic_u.vss", "Rectangle"));
            }
            layout.AddEdge("e0", nodes[0], nodes[1], "", VA.Models.ConnectorType.Straight);
            layout.AddEdge("e1", nodes[0], nodes[2], "", VA.Models.ConnectorType.Straight);
            layout.AddEdge("e2", nodes[2], nodes[3], "", VA.Models.ConnectorType.Straight);
            return layout;
        }

        // Draws a one-layout document and returns the margin around the node shapes on its page.
        // layout_border: the layout's PageBorderWidth, or null to leave it unset.
        private double draw_one_page(double? layout_border, double document_border)
        {
            var layout = make_layout();
            if (layout_border.HasValue)
            {
                layout.LayoutOptions.PageBorderWidth = new VA.Core.Size(layout_border.Value, layout_border.Value);
            }
            var dgdoc = new VADG.DirectedGraphDocument();
            dgdoc.Layouts.Add(layout);
            dgdoc.BorderSize = new VA.Core.Size(document_border, document_border);
            return this.draw_and_measure(dgdoc)[0];
        }

        // Draws the document in a new Visio document, returns the horizontal margin around the node shapes on each page,
        // and closes the Visio document.
        private List<double> draw_and_measure(VADG.DirectedGraphDocument dgdoc)
        {
            var client = this.GetScriptingClient();
            client.Model.DrawDirectedGraphDocument(dgdoc, new VADG.DirectedGraphStyling());

            var doc = this.GetVisioApplication().ActiveDocument;
            var margins = new List<double>();
            var pages = doc.Pages;
            for (int i = 1; i <= pages.Count; i++)
            {
                margins.Add(measure_margin(pages[i]));
            }
            doc.Close(true);
            return margins;
        }

        private double render_directly(double? border)
        {
            var layout = make_layout();
            if (border.HasValue)
            {
                layout.LayoutOptions.PageBorderWidth = new VA.Core.Size(border.Value, border.Value);
            }

            var visapp = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = visapp.ActivePage;

            var renderer = new VADG.MsaglRenderer();
            renderer.LayoutOptions = layout.LayoutOptions;
            renderer.Render(page, layout);

            double margin = measure_margin(page);
            doc.Close(true);
            return margin;
        }

        // Half of the page width left over after the node shapes (the 2D shapes; connectors are 1D and are skipped).
        private static double measure_margin(Microsoft.Office.Interop.Visio.Page page)
        {
            double page_width = page.PageSheet.CellsU["PageWidth"].ResultIU;
            double min_x = double.MaxValue;
            double max_x = double.MinValue;
            foreach (Microsoft.Office.Interop.Visio.Shape shape in page.Shapes)
            {
                if (shape.OneD != 0)
                {
                    continue;
                }
                double pin_x = shape.CellsU["PinX"].ResultIU;
                double width = shape.CellsU["Width"].ResultIU;
                min_x = System.Math.Min(min_x, pin_x - width / 2.0);
                max_x = System.Math.Max(max_x, pin_x + width / 2.0);
            }
            return (page_width - (max_x - min_x)) / 2.0;
        }
    }
}

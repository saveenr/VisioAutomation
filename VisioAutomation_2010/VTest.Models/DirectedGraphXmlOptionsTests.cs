using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SXL = System.Xml.Linq;
using VA = VisioAutomation;
using VADG = VisioAutomation.Models.Layouts.DirectedGraph;

namespace VTest.Models
{
    // Tests for the optional parts of the directed graph XML format added for issue #225:
    // <documentoptions>, page border attributes, shape width/height, <hyperlink>, <cells>,
    // typed <customprop>, and per-connector connectortype, <cells> and <customprop>.
    [MUT.TestClass]
    public class DirectedGraphXmlOptionsTests : Framework.VTest
    {
        private string build_xml(
            string root_children = "",
            string render_attrs = "",
            string shape_attrs = "",
            string shape_children = "",
            string connector_attrs = "",
            string connector_children = "")
        {
            return
                "<directedgraph>" + root_children +
                "<page>" +
                "<renderoptions usedynamicconnectors='true' scalingfactor='20' " + render_attrs + " />" +
                "<shapes>" +
                "<shape id='n1' label='A' stencil='basic_u.vss' master='Rectangle' " + shape_attrs + ">" + shape_children + "</shape>" +
                "<shape id='n2' label='B' stencil='basic_u.vss' master='Rectangle' />" +
                "</shapes>" +
                "<connectors>" +
                "<connector id='c1' from='n1' to='n2' label='' " + connector_attrs + ">" + connector_children + "</connector>" +
                "</connectors>" +
                "</page>" +
                "</directedgraph>";
        }

        private VADG.DirectedGraphDocument load(string xml)
        {
            var client = this.GetScriptingClient();
            return client.Model.LoadDirectedGraphFromXml(SXL.XDocument.Parse(xml));
        }

        // ---- <documentoptions> ----

        [MUT.TestMethod]
        public void Loader_DocumentOptions_TemplateAndBorderFromXml()
        {
            var dg = this.load(this.build_xml(root_children: "<documentoptions template='basflo_u.vst' borderwidth='2' borderheight='3' />"));
            MUT.Assert.AreEqual("basflo_u.vst", dg.Template);
            MUT.Assert.AreEqual(2.0, dg.BorderSize.Width);
            MUT.Assert.AreEqual(3.0, dg.BorderSize.Height);
        }

        [MUT.TestMethod]
        public void Loader_DocumentOptions_AbsentKeepsDefaults()
        {
            var dg = this.load(this.build_xml());
            MUT.Assert.IsNull(dg.Template);
            MUT.Assert.AreEqual(1.0, dg.BorderSize.Width);
            MUT.Assert.AreEqual(1.0, dg.BorderSize.Height);
        }

        [MUT.TestMethod]
        public void Loader_DocumentOptions_OneBorderAttributeKeepsTheOtherDefault()
        {
            var dg = this.load(this.build_xml(root_children: "<documentoptions borderwidth='4' />"));
            MUT.Assert.AreEqual(4.0, dg.BorderSize.Width);
            MUT.Assert.AreEqual(1.0, dg.BorderSize.Height);
        }

        // ---- <renderoptions> page border ----

        [MUT.TestMethod]
        public void Loader_PageBorder_FromXml()
        {
            var dg = this.load(this.build_xml(render_attrs: "pageborderwidth='0.25' pageborderheight='0.75'"));
            var border = dg.Layouts[0].LayoutOptions.PageBorderWidth;
            MUT.Assert.AreEqual(0.25, border.Width);
            MUT.Assert.AreEqual(0.75, border.Height);
        }

        [MUT.TestMethod]
        public void Loader_PageBorder_AbsentKeepsDefault()
        {
            var dg = this.load(this.build_xml());
            var border = dg.Layouts[0].LayoutOptions.PageBorderWidth;
            MUT.Assert.AreEqual(0.5, border.Width);
            MUT.Assert.AreEqual(0.5, border.Height);
        }

        // ---- <shape> ----

        [MUT.TestMethod]
        public void Loader_ShapeSize_FromXml()
        {
            var dg = this.load(this.build_xml(shape_attrs: "width='3' height='2'"));
            var size = dg.Layouts[0].Nodes["n1"].Size;
            MUT.Assert.IsTrue(size.HasValue);
            MUT.Assert.AreEqual(3.0, size.Value.Width);
            MUT.Assert.AreEqual(2.0, size.Value.Height);
            MUT.Assert.IsFalse(dg.Layouts[0].Nodes["n2"].Size.HasValue);
        }

        [MUT.TestMethod]
        public void Loader_ShapeSize_WidthWithoutHeight_ThrowsArgumentException()
        {
            MUT.Assert.ThrowsExactly<System.ArgumentException>(() => this.load(this.build_xml(shape_attrs: "width='3'")));
        }

        [MUT.TestMethod]
        public void Loader_ShapeHyperlinks_FromXml()
        {
            var dg = this.load(this.build_xml(shape_children:
                "<hyperlink name='Docs' address='https://example.com/docs' subaddress='top' description='the docs' />" +
                "<hyperlink name='Repo' address='https://example.com/repo' />"));
            var links = dg.Layouts[0].Nodes["n1"].Hyperlinks;
            MUT.Assert.AreEqual(2, links.Count);
            MUT.Assert.AreEqual("Docs", links[0].Name);
            MUT.Assert.AreEqual("https://example.com/docs", links[0].Address);
            MUT.Assert.AreEqual("top", links[0].SubAddress);
            MUT.Assert.AreEqual("the docs", links[0].Description);
            MUT.Assert.AreEqual("Repo", links[1].Name);
        }

        [MUT.TestMethod]
        public void Loader_ShapeHyperlinks_UrlAttributeStaysFirst()
        {
            var dg = this.load(this.build_xml(
                shape_attrs: "url='https://example.com/first'",
                shape_children: "<hyperlink name='Second' address='https://example.com/second' />"));
            var links = dg.Layouts[0].Nodes["n1"].Hyperlinks;
            MUT.Assert.AreEqual(2, links.Count);
            MUT.Assert.AreEqual("https://example.com/first", links[0].Address);
            MUT.Assert.AreEqual("Second", links[1].Name);
        }

        [MUT.TestMethod]
        public void Loader_ShapeHyperlinks_UrlAttributeAloneLeavesHyperlinksNull()
        {
            // existing behavior: a url attribute alone is turned into a hyperlink by the renderer
            var dg = this.load(this.build_xml(shape_attrs: "url='https://example.com/only'"));
            MUT.Assert.IsNull(dg.Layouts[0].Nodes["n1"].Hyperlinks);
            MUT.Assert.AreEqual("https://example.com/only", dg.Layouts[0].Nodes["n1"].Url);
        }

        [MUT.TestMethod]
        public void Loader_ShapeCells_FromXml()
        {
            var dg = this.load(this.build_xml(shape_children:
                "<cells><cell name='FillForeground' value='RGB(255,0,0)' /><cell name='charsize' value='14 pt' /></cells>"));
            var cells = dg.Layouts[0].Nodes["n1"].Cells;
            MUT.Assert.IsNotNull(cells);
            MUT.Assert.AreEqual("RGB(255,0,0)", cells.FillForeground.Value);
            MUT.Assert.AreEqual("14 pt", cells.CharSize.Value, "cell names are matched without regard to case");
            MUT.Assert.IsNull(dg.Layouts[0].Nodes["n2"].Cells, "a shape without <cells> keeps null Cells");
        }

        [MUT.TestMethod]
        public void Loader_ShapeCells_UnknownCellName_ThrowsArgumentException()
        {
            var ex = MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => this.load(this.build_xml(shape_children: "<cells><cell name='NotACell' value='1' /></cells>")));
            MUT.StringAssert.Contains(ex.Message, "NotACell");
        }

        [MUT.TestMethod]
        public void Loader_ShapeCustomProperty_TypedFromXml()
        {
            var dg = this.load(this.build_xml(shape_children:
                "<customprop name='s' value='text' />" +
                "<customprop name='n' value='2.5' type='number' label='Count' />" +
                "<customprop name='b' value='true' type='boolean' />"));
            var props = dg.Layouts[0].Nodes["n1"].CustomProperties;

            MUT.Assert.AreEqual("\"text\"", props["s"].Formula.Value);
            MUT.Assert.AreEqual("0", props["s"].Type.Value);

            MUT.Assert.AreEqual("2.5", props["n"].Formula.Value);
            MUT.Assert.AreEqual("2", props["n"].Type.Value);
            MUT.Assert.AreEqual("\"Count\"", props["n"].Label.Value);

            MUT.Assert.AreEqual("TRUE", props["b"].Formula.Value);
            MUT.Assert.AreEqual("3", props["b"].Type.Value);
        }

        [MUT.TestMethod]
        public void Loader_ShapeCustomProperty_UnknownType_ThrowsArgumentException()
        {
            MUT.Assert.ThrowsExactly<System.ArgumentException>(
                () => this.load(this.build_xml(shape_children: "<customprop name='x' value='1' type='widget' />")));
        }

        // ---- <connector> ----

        [MUT.TestMethod]
        public void Loader_ConnectorType_PerEdgeAttributeOverridesThePageSetting()
        {
            var dg = this.load(this.build_xml(render_attrs: "connectortype='Straight'", connector_attrs: "connectortype='RightAngle'"));
            MUT.Assert.AreEqual(VA.Models.ConnectorType.RightAngle, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorType_WithoutAnEdgeAttributeUsesThePageSetting()
        {
            var dg = this.load(this.build_xml(render_attrs: "connectortype='Straight'"));
            MUT.Assert.AreEqual(VA.Models.ConnectorType.Straight, dg.Layouts[0].Edges["c1"].ConnectorType);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorCells_WinOverTheColorAndWeightDefaults()
        {
            var dg = this.load(this.build_xml(
                connector_attrs: "color='#ff0000'",
                connector_children: "<cells><cell name='LineColor' value='RGB(0,0,255)' /><cell name='LinePattern' value='2' /></cells>"));
            var cells = dg.Layouts[0].Edges["c1"].Cells;
            MUT.Assert.AreEqual("RGB(0,0,255)", cells.LineColor.Value, "an explicit cell beats the color attribute");
            MUT.Assert.AreEqual("2", cells.LinePattern.Value);
            MUT.Assert.IsNotNull(cells.LineWeight.Value, "the weight default is still set");
        }

        [MUT.TestMethod]
        public void Loader_ConnectorCustomProperties_FromXml()
        {
            var dg = this.load(this.build_xml(connector_children: "<customprop name='owner' value='ops' />"));
            var props = dg.Layouts[0].Edges["c1"].CustomProperties;
            MUT.Assert.IsNotNull(props);
            MUT.Assert.AreEqual("\"ops\"", props["owner"].Formula.Value);
        }

        [MUT.TestMethod]
        public void Loader_ConnectorWithoutCustomProperties_LeavesThemNull()
        {
            var dg = this.load(this.build_xml());
            MUT.Assert.IsNull(dg.Layouts[0].Edges["c1"].CustomProperties);
        }

        // ---- drawing ----

        private VADG.DirectedGraphDocument draw(string xml)
        {
            var client = this.GetScriptingClient();
            var dgdoc = client.Model.LoadDirectedGraphFromXml(SXL.XDocument.Parse(xml));
            client.Model.DrawDirectedGraphDocument(dgdoc, new VADG.DirectedGraphStyling());
            return dgdoc;
        }

        [MUT.TestMethod]
        public void Draw_ShapeSizeAndCellsFromXml_ReachTheVisioShape()
        {
            var dgdoc = this.draw(this.build_xml(
                shape_attrs: "width='3' height='2'",
                shape_children: "<cells><cell name='FillForeground' value='RGB(255,0,0)' /></cells>"));
            var doc = this.GetVisioApplication().ActiveDocument;

            var shape = dgdoc.Layouts[0].Nodes["n1"].VisioShape;
            MUT.Assert.AreEqual(3.0, shape.CellsU["Width"].ResultIU, 1e-6);
            MUT.Assert.AreEqual(2.0, shape.CellsU["Height"].ResultIU, 1e-6);
            MUT.Assert.AreEqual("RGB(255,0,0)", shape.CellsU["FillForegnd"].FormulaU.Replace(" ", "").ToUpperInvariant());

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Draw_ShapeHyperlinksFromXml_AreAddedToTheShape()
        {
            var dgdoc = this.draw(this.build_xml(
                shape_attrs: "url='https://example.com/first'",
                shape_children: "<hyperlink name='Second' address='https://example.com/second' />"));
            var doc = this.GetVisioApplication().ActiveDocument;

            var shape = dgdoc.Layouts[0].Nodes["n1"].VisioShape;
            MUT.Assert.AreEqual(2, shape.Hyperlinks.Count);

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Draw_ConnectorCustomPropertiesFromXml_AreSetOnDynamicConnectors()
        {
            var dgdoc = this.draw(this.build_xml(connector_children: "<customprop name='owner' value='ops' />"));
            var doc = this.GetVisioApplication().ActiveDocument;

            var shape = dgdoc.Layouts[0].Edges["c1"].VisioShape;
            var props = VA.Shapes.CustomPropertyHelper.GetDictionary(shape, VA.Core.CellValueType.Formula);
            MUT.Assert.IsTrue(props.ContainsKey("owner"));
            MUT.Assert.AreEqual("\"ops\"", props["owner"].Formula.Value);

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Draw_ConnectorCustomPropertiesFromXml_AreSetOnBezierConnectors()
        {
            var xml = this.build_xml(connector_children: "<customprop name='owner' value='ops' />").Replace("usedynamicconnectors='true'", "usedynamicconnectors='false'");
            var dgdoc = this.draw(xml);
            var doc = this.GetVisioApplication().ActiveDocument;

            var shape = dgdoc.Layouts[0].Edges["c1"].VisioShape;
            var props = VA.Shapes.CustomPropertyHelper.GetDictionary(shape, VA.Core.CellValueType.Formula);
            MUT.Assert.IsTrue(props.ContainsKey("owner"));

            doc.Close(true);
        }

        [MUT.TestMethod]
        public void Draw_DocumentBorderFromXml_ChangesThePageSize()
        {
            double default_width = this.draw_and_get_page_width(this.build_xml());
            double wide_width = this.draw_and_get_page_width(this.build_xml(root_children: "<documentoptions borderwidth='4' borderheight='4' />"));

            MUT.Assert.IsTrue(wide_width > default_width + 3.0,
                string.Format("expected a larger border to widen the page: default {0}, wide {1}", default_width, wide_width));
        }

        private double draw_and_get_page_width(string xml)
        {
            this.draw(xml);
            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            double width = app.ActivePage.PageSheet.CellsU["PageWidth"].ResultIU;
            doc.Close(true);
            return width;
        }
    }
}

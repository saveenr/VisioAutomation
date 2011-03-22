using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class DOM_Draw_Tests : VisioAutomationTest
    {

        public int get_doc_count( IVisio.Application app)
        {
            // get the number of actual drawings, not including templates, stencils, etc.
            var documents = app.Documents;
            return documents.AsEnumerable()
                .Where( doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing)
                .Count();
        }

        [TestMethod]
        public void Empty_DOM_Rendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();


            var doc1 = new VA.DOM.Document();


            var doc = this.GetNewDoc();
            doc1.Render(app.ActivePage);

            Assert.AreEqual(0,app.ActivePage.Shapes.Count);
            
            app.ActiveDocument.Close( true );
        }

        [TestMethod]
        public void Empty_DOM_Page_Size()
        {

            // A DOM document with 1 pages rendered to a document with 1 page should ????
            var app = this.GetVisioApplication();

            var doc1 = new VA.DOM.Document();
            doc1.PageSettings.Size = new VA.Drawing.Size(5,5);

            var visdoc = this.GetNewDoc();
            Assert.AreEqual(1, visdoc.Pages.Count);

            doc1.Render(app.ActivePage);

            Assert.AreEqual(1, visdoc.Pages.Count);
            TestUtil.AreEqual(5, 5, app.ActivePage.GetSize(), 0.005);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Draw_Red_Rectangle_With_Text()
        {
            // Create the doc
            var vdoc = new VA.DOM.Document();
            vdoc.PageSettings.Size = new VA.Drawing.Size(10,10);
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = "HELLO WORLD";
            vrect1.ShapeCells.FillForegnd = VA.Convert.ColorToFormulaRGB(0xff0000);
            vdoc.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            vdoc.Render(app.ActivePage);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);

            app.ActiveDocument.Close(true);
        }


        [TestMethod]
        public void MarkupText1()
        {
            var vdoc = new VA.DOM.Document();
            vdoc.PageSettings.Size = new VA.Drawing.Size(10,10);

            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vdoc.Shapes.Add(vrect1);

            string text =
                @"
<text>
    Normal text
        <br/>

    <text font=""Courier New"">
        Courier new at default size
        <br/>

        <text size=""20"">
            Now at 20pt 
        <br/>
                <text bold=""1"">and this text is bold</text>
                <text italic=""1"">and this text is italic</text>
        </text>
    </text>

</text> ";

            var text_markup = VA.Text.Markup.TextElement.FromXml(text, false);
            vrect1.TextElement = text_markup;

            vrect1.SetCustomProperty("FOO1", "bar");
            vrect1.SetCustomProperty("FOO2", "\"bar\"");
            vrect1.SetCustomProperty("FOO3", "\"\"bar\"\"");

            var app = this.GetVisioApplication();
            var documents = app.Documents;
            var visdoc = this.GetNewDoc();
            vdoc.Render(app.ActivePage);

            var activepagesize = app.ActivePage.GetSize();
            Assert.AreEqual(10, activepagesize.Width);
            Assert.AreEqual(10, activepagesize.Height);

            visdoc.Close(true);
        }

        [TestMethod]
        public void Set_Custom_Props()
        {
            //Draws a simple red square

            // Create the doc
            var vdoc = new VA.DOM.Document();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = "HELLO WORLD";

            var cp1 = vrect1.SetCustomProperty("FOO", "FOOVALUE");
            cp1.Label = "Foo Label";
            var cp2 = vrect1.SetCustomProperty("BAR", "BARVALUE");
            cp2.Label = "Bar Label";

            vdoc.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            vdoc.Render(app.ActivePage);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(vrect1.VisioShape, "FOO"));
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(vrect1.VisioShape, "BAR"));

            doc.Close(true);
        }

        [TestMethod]
        public void MarkupText22()
        {
            string text =
                @"
<text>
    Normal text
        <br/>

    <text font=""Courier New"">
        Courier new at default size
        <br/>

        <text size=""20"">
            Now at 20pt 
        <br/>
                <text bold=""1"">and this text is bold</text>
                <text italic=""1"">and this text is italic</text>
        </text>
    </text>

</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);

            Assert.AreEqual(1, root_el.Children.Count);
            var root_elements = root_el.Elements.ToList();

            var n0 = root_el.Children[0];
            Assert.AreEqual(3, n0.Children.Count);
            var n1 = n0.Children[0];
            var n2 = n0.Children[1];
            var n3 = n0.Children[2];
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n1.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n2.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n3.NodeType);

            Assert.AreEqual("Normal text", n1.GetInnerText());
            Assert.AreEqual("\n", n2.GetInnerText());
            Assert.AreEqual(3, n3.Children.Count);

            var n4 = n3.Children[0];
            var n5 = n3.Children[1];
            var n6 = n3.Children[2];
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n4.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n5.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n6.NodeType);

            Assert.AreEqual("Courier new at default size", n4.GetInnerText());
            Assert.AreEqual("\n", n5.GetInnerText());
            Assert.AreEqual("Now at 20pt\nand this text is boldand this text is italic", n6.GetInnerText());
        }

        [TestMethod]
        public void Markup_Simple_Plain()
        {
            string text =
                @"
<text>
    Normal Text
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Bold()
        {
            string text =
                @"
<text bold=""1"">
    Bold Text
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Italic()
        {
            string text =
                @"
<text italic=""1"">
    Italic Text
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Font()
        {
            string text =
                @"
<text font=""Impact"">
    Italic Text
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Render_Markup_Simple_Font_Multiple()
        {
            string text =
                @"
<text font=""Impact"" color=""#ff0000"">
    Impact font red
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Overlap_Multiple()
        {
            string text =
                @"
<text>
    plain
    <text italic=""1"" >
        italic
        <text bold=""1"" >
            bold
        </text>
    </text>
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            root_el.SetShapeText(s0);
            page1.Delete(0);
        }
    }
}
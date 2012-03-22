using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.CustomProperties;
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
            AssertVA.AreEqual(5, 5, app.ActivePage.GetSize(), 0.005);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Draw_Red_Rectangle_With_Text()
        {
            // Create the doc
            var vdoc = new VA.DOM.Document();
            vdoc.PageSettings.Size = new VA.Drawing.Size(10,10);
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");
            vrect1.Cells.FillForegnd = VA.Convert.ColorToFormulaRGB(0xff0000);
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
        public void Draw_DropShapes()
        {
            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var stencil = app.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];

            // Create the doc
            var vdoc = new VA.DOM.Document();
            
            vdoc.DrawRectangle(0, 0, 1, 1);
            vdoc.Drop(rectmaster, 3, 3);

            vdoc.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Set_Custom_Props()
        {
            //Draws a simple red square

            // Create the doc
            var vdoc = new VA.DOM.Document();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");

            vrect1.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();

            var cp1 = new VA.CustomProperties.CustomPropertyCells();
            cp1.Value = "FOOVALUE";
            cp1.Label = "Foo Label";

            var cp2 = new VA.CustomProperties.CustomPropertyCells();
            cp2.Value = "BARVALUE";
            cp2.Label = "Bar Label";

            vrect1.CustomProperties["FOO"] = cp1;
            vrect1.CustomProperties["BAR"] = cp2;

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
        public void Markup_Simple_Plain()
        {
            var m = new VA.Text.Markup.TextElement("Normal Text");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Bold()
        {
            var m = new VA.Text.Markup.TextElement("Normal Text");
            m.CharacterFormat.CharStyle = VA.Text.CharStyle.Bold;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Italic()
        {
            var m = new VA.Text.Markup.TextElement("Normal Text");
            m.CharacterFormat.CharStyle = VA.Text.CharStyle.Italic;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Font()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));

            var impact = page1.Document.Fonts["Impact"];
            var m = new VA.Text.Markup.TextElement("Normal Text");
            m.CharacterFormat.FontID = impact.ID;
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Render_Markup_Simple_Font_Multiple()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var impact = page1.Document.Fonts["Impact"];
            var m = new VA.Text.Markup.TextElement("Normal Text");
            m.CharacterFormat.FontID = impact.ID;
            m.CharacterFormat.Color = new VA.Drawing.ColorRGB(0xff0000);
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);
            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Overlap_Multiple()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var segoeui = page1.Document.Fonts["Segoe UI"];

            var t1 = new VA.Text.Markup.TextElement("Normal Text");
            t1.CharacterFormat.FontID = segoeui.ID;
            var t2 = t1.AppendElement("Italic");
            t2.CharacterFormat.CharStyle = VA.Text.CharStyle.Italic;

            var t3 = t2.AppendElement("Italic");
            t3.CharacterFormat.CharStyle = VA.Text.CharStyle.Bold;

            var t4 = t2.AppendElement("Bold Italic");
            t4.CharacterFormat.CharStyle = VA.Text.CharStyle.Bold | VA.Text.CharStyle.Italic;

            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            t1.SetText(s0);
            page1.Delete(0);
        }
    }
}
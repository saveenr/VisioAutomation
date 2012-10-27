using System.Collections.Generic;
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
            var drawings = documents.AsEnumerable()
                .Where(doc => doc.Type == IVisio.VisDocumentTypes.visTypeDrawing);
            return drawings.Count();
        }

        [TestMethod]
        public void Empty_DOM_Rendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();

            var dompage = new VA.DOM.Page();
            var doc = this.GetNewDoc();
            dompage.Render(app.ActiveDocument);
            Assert.AreEqual(0,app.ActivePage.Shapes.Count);
            app.ActiveDocument.Close( true );
        }

        [TestMethod]
        public void Render_Page_To_Document()
        {
            // Rendering a dom page to a document should create a new page
            var app = this.GetVisioApplication();
            var dompage = new VA.DOM.Page();
            var visdoc = this.GetNewDoc();
            Assert.AreEqual(1, visdoc.Pages.Count);
            var page = dompage.Render(app.ActiveDocument);  
            Assert.AreEqual(2, visdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Render_Document_To_App()
        {
            // Rendering a dom document to an appliction instance should create a new document
            var app = this.GetVisioApplication();
            var domdoc = new VA.DOM.Document();
            var dompage = new VA.DOM.Page();
            domdoc.Pages.Add(dompage);
            int old_count = app.Documents.Count;
            var newdoc = domdoc.Render(app);
            Assert.AreEqual(old_count + 1, app.Documents.Count);
            Assert.AreEqual(1, newdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Draw_Red_Rectangle_With_Text()
        {
            // Create the doc
            var dompage = new VA.DOM.Page();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");
            vrect1.Cells.FillForegnd = "rgb(255,0,0)";
            dompage.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            app.ActivePage.SetSize(new VA.Drawing.Size(10, 10));
            var page = dompage.Render(app.ActiveDocument);

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
            var domshapescol = new VA.DOM.ShapeList();
            
            domshapescol.DrawRectangle(0, 0, 1, 1);
            domshapescol.Drop(rectmaster, 3, 3);

            domshapescol.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Set_Custom_Props()
        {
            // Create the doc
            var domshapescol = new VA.DOM.ShapeList();
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

            domshapescol.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            domshapescol.Render(app.ActivePage);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(vrect1.VisioShape, "FOO"));
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(vrect1.VisioShape, "BAR"));

            doc.Close(true);
        }
    }
}
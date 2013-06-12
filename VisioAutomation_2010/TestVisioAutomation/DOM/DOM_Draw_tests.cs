using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DOM;
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

        public void DOM_Render_Scenarios()
        {
            this.Empty_DOM_Rendering();
            this.Render_Page_To_Document();
            this.Render_Document_To_App();
        }

        public void Empty_DOM_Rendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();

            var page_node = new VA.DOM.Page();
            var doc = this.GetNewDoc();
            page_node.Render(app.ActiveDocument);
            Assert.AreEqual(0,app.ActivePage.Shapes.Count);
            app.ActiveDocument.Close( true );
        }

        public void Render_Page_To_Document()
        {
            // Rendering a dom page to a document should create a new page
            var app = this.GetVisioApplication();
            var page_node = new VA.DOM.Page();
            var visdoc = this.GetNewDoc();
            Assert.AreEqual(1, visdoc.Pages.Count);
            var page = page_node.Render(app.ActiveDocument);  
            Assert.AreEqual(2, visdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        public void Render_Document_To_App()
        {
            // Rendering a dom document to an appliction instance should create a new document
            var app = this.GetVisioApplication();
            var doc_node = new VA.DOM.Document();
            var page_node = new VA.DOM.Page();
            doc_node.Pages.Add(page_node);
            int old_count = app.Documents.Count;
            var newdoc = doc_node.Render(app);
            Assert.AreEqual(old_count + 1, app.Documents.Count);
            Assert.AreEqual(1, newdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void BasicDOMDrawing()
        {
            this.DrawSimpleShape();
            this.DropShapes();
            this.SetCustomProperties();
            this.DrawOrgChart();
        }

        public void DrawSimpleShape()
        {
            // Create the doc
            var page_node = new VA.DOM.Page();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");
            vrect1.Cells.FillForegnd = "rgb(255,0,0)";
            page_node.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            VisioAutomationTest.SetPageSize(app.ActivePage, new VA.Drawing.Size(10, 10));
            var page = page_node.Render(app.ActiveDocument);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);

            app.ActiveDocument.Close(true);
        }

        public void DropShapes()
        {
            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var stencil = app.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];


            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();
            
            shape_nodes.DrawRectangle(0, 0, 1, 1);
            shape_nodes.Drop(rectmaster, 3, 3);

            shape_nodes.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        public void SetCustomProperties()
        {
            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();
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

            shape_nodes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            shape_nodes.Render(app.ActivePage);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "FOO"));
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "BAR"));

            doc.Close(true);
        }

        public void DrawOrgChart()
        {
            // How to draw using a Template instead of a doc and a stencil
            const string orgchart_vst = "orgch_u.vst";

            var app = this.GetVisioApplication();
            var doc_node = new VA.DOM.Document( orgchart_vst , IVisio.VisMeasurementSystem.visMSUS );
            var page_node = new VA.DOM.Page();
            doc_node.Pages.Add(page_node);

            // Have to be smart about selecting the right master with Visio 2013
            int vis_ver = int.Parse(app.Version.Split( new char[]{'.'} )[0]);
            string position_master_name = vis_ver >= 15 ? "Position Belt" : "Position";

            var s1 = new Shape(position_master_name, null, new VA.Drawing.Point(3, 4));
            page_node.Shapes.Add( s1 );
            var doc = doc_node.Render(app);

            doc.Close(true);
        }
    }
}
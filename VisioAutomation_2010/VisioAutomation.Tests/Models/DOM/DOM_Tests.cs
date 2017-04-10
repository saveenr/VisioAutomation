using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using VisioAutomation.Models.Dom;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation_Tests.Models.Dom
{
    [TestClass]
    public class Dom_Tests : VisioAutomationTest
    {
        public string basic_u_vss = "basic_u.vss";
        public string connec_u_vss = "connec_u.vss";
        public string rectangle = "Rectangle";
        public string dynamicconnector = "Dynamic Connector";
        private VisioAutomation.Drawing.Size pagesize;

        [TestMethod]
        public void Dom_EmptyRendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();

            var page_node = new Page();
            var doc = this.GetNewDoc();
            page_node.Render(app.ActiveDocument);
            Assert.AreEqual(0, app.ActivePage.Shapes.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Dom_RenderPageToDocument()
        {
            // Rendering a dom page to a document should create a new page
            var app = this.GetVisioApplication();
            var page_node = new Page();
            var visdoc = this.GetNewDoc();
            Assert.AreEqual(1, visdoc.Pages.Count);
            var page = page_node.Render(app.ActiveDocument);
            Assert.AreEqual(2, visdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Dom_RenderDocumentToApplication()
        {
            // Rendering a dom document to an appliction instance should create a new document
            var app = this.GetVisioApplication();
            var doc_node = new Document();
            var page_node = new Page();
            doc_node.Pages.Add(page_node);
            int old_count = app.Documents.Count;
            var newdoc = doc_node.Render(app);
            Assert.AreEqual(old_count + 1, app.Documents.Count);
            Assert.AreEqual(1, newdoc.Pages.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Dom_DrawSimpleShape()
        {
            // Create the doc
            var page_node = new Page();
            var vrect1 = new Rectangle(1, 1, 9, 9);
            vrect1.Text = new VisioAutomation.Models.Text.TextElement("HELLO WORLD");
            vrect1.Cells.FillForeground = "rgb(255,0,0)";
            page_node.Shapes.Add(vrect1);

            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            this.pagesize = new VA.Drawing.Size(10, 10);
            VisioAutomationTest.SetPageSize(app.ActivePage, this.pagesize);
            var page = page_node.Render(app.ActiveDocument);

            // Verify
            Assert.IsNotNull(vrect1.VisioShape);
            Assert.AreEqual("HELLO WORLD", vrect1.VisioShape.Text);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Dom_DropShapes()
        {
            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var stencil = app.Documents.OpenStencil(this.basic_u_vss);
            var rectmaster = stencil.Masters[this.rectangle];


            // Create the doc
            var shape_nodes = new ShapeList();

            shape_nodes.DrawRectangle(0, 0, 1, 1);
            shape_nodes.Drop(rectmaster, 3, 3);

            shape_nodes.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void Dom_CustomProperties()
        {
            // Create the doc
            var shape_nodes = new ShapeList();
            var vrect1 = new Rectangle(1, 1, 9, 9);
            vrect1.Text = new VisioAutomation.Models.Text.TextElement("HELLO WORLD");

            vrect1.CustomProperties = new VisioAutomation.Shapes.CustomProperties.CustomPropertyDictionary();

            var cp1 = new VA.Shapes.CustomProperties.CustomPropertyCells();
            cp1.Value = "FOOVALUE";
            cp1.Label = "Foo Label";

            var cp2 = new VA.Shapes.CustomProperties.CustomPropertyCells();
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
            Assert.IsTrue(VA.Shapes.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "FOO"));
            Assert.IsTrue(VA.Shapes.CustomProperties.CustomPropertyHelper.Contains(vrect1.VisioShape, "BAR"));

            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DrawOrgChart()
        {
            var app = this.GetVisioApplication();
            var vis_ver = VA.Application.ApplicationHelper.GetVersion(app);

            // How to draw using a Template instead of a doc and a stencil
            string orgchart_vst = "orgchart.vst";
            string position_master_name = vis_ver.Major >= 15 ? "Position Belt" : "Position";

            var doc_node = new Document(orgchart_vst, IVisio.VisMeasurementSystem.visMSUS);
            var page_node = new Page();
            doc_node.Pages.Add(page_node);

            // Have to be smart about selecting the right master with Visio 2013

            var s1 = new Shape(position_master_name, null, new VA.Drawing.Point(3, 8));
            page_node.Shapes.Add(s1);

            var s2 = new Shape(position_master_name, null, new VA.Drawing.Point(0, 4));
            page_node.Shapes.Add(s2);

            var s3 = new Shape(position_master_name, null, new VA.Drawing.Point(6, 4));
            page_node.Shapes.Add(s3);

            page_node.Shapes.Connect(this.dynamicconnector, this.connec_u_vss, s1, s2);
            page_node.Shapes.Connect(this.dynamicconnector, this.connec_u_vss, s1, s3);

            var doc = doc_node.Render(app);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DrawEmpty()
        {
            // Verify that an empty DOM page can be created and rendered
            var doc = this.GetNewDoc();
            var page_node = new Page();
            page_node.Size = new VA.Drawing.Size(5, 5);
            var page = page_node.Render(doc);

            Assert.AreEqual(0, page.Shapes.Count);
            var actual_page_size = VisioAutomationTest.GetPageSize(page);
            var expected_page_size = new VA.Drawing.Size(5, 5);
            Assert.AreEqual(expected_page_size, actual_page_size);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DrawLine()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();
            var line_node_0 = page_node.Shapes.DrawLine(1, 1, 3, 3);
            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, line_node_0.VisioShapeID);
            Assert.IsNotNull(line_node_0.VisioShape);
            Assert.AreEqual(2.0, line_node_0.VisioShape.CellsU["PinX"].Result[IVisio.VisUnitCodes.visNumber]);
            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DrawBezier()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();
            var bez_node_0 = page_node.Shapes.DrawBezier(new double[] { 1, 2, 3, 3, 6, 3, 3, 4 });

            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, bez_node_0.VisioShapeID);
            Assert.IsNotNull(bez_node_0.VisioShape);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DropMaster()
        {

            var doc = this.GetNewDoc();
            var page_node = new Page();
            var stencil = doc.Application.Documents.OpenStencil(this.basic_u_vss);
            var master1 = stencil.Masters[this.rectangle];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(this.rectangle, this.basic_u_vss, 5, 5);

            var page = page_node.Render(doc);

            Assert.AreEqual(2, page.Shapes.Count);

            // Verify that the shapes created both have IDs and shape objects associated with them
            Assert.AreNotEqual(0, master_node_0.VisioShapeID);
            Assert.AreNotEqual(0, master_node_1.VisioShapeID);
            Assert.IsNotNull(master_node_0.VisioShape);
            Assert.IsNotNull(master_node_1.VisioShape);
            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_FormatShape()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();
            var stencil = doc.Application.Documents.OpenStencil(this.basic_u_vss);
            var master1 = stencil.Masters[this.rectangle];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var bez_node_0 = page_node.Shapes.DrawBezier(new double[] { 1, 2, 3, 3, 6, 3, 3, 4 });
            var line_node_0 = page_node.Shapes.DrawLine(1, 1, 3, 3);

            master_node_0.Cells.LineWeight = 0.1;
            bez_node_0.Cells.LineWeight = 0.3;
            line_node_0.Cells.LineWeight = 0.5;

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);
            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_ConnectShapes()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();

            var basic_stencil = doc.Application.Documents.OpenStencil(this.basic_u_vss);
            var basic_masters = basic_stencil.Masters;
            var connectors_stencil = doc.Application.Documents.OpenStencil(this.connec_u_vss);
            var connectors_masters = connectors_stencil.Masters;

            var master1 = basic_masters[this.rectangle];
            var master2 = connectors_masters[this.dynamicconnector];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(master1, 6, 5);
            var dc = page_node.Shapes.Connect(master2, master_node_0, master_node_1);

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_ConnectShapes2()
        {
            // Deferred means that the stencils (and thus masters) are loaded when rendering
            // and are no loaded by the caller before Render() is called

            var doc = this.GetNewDoc();
            var page_node = new Page();
            var master_node_0 = page_node.Shapes.Drop(this.rectangle, this.basic_u_vss, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(this.rectangle, this.basic_u_vss, 6, 5);
            var dc = page_node.Shapes.Connect(this.dynamicconnector, this.connec_u_vss, master_node_0, master_node_1);
            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_VerifyThatUnknownMastersAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();
            var master_node_0 = page_node.Shapes.Drop("XXX", this.basic_u_vss, 3, 3);

            IVisio.Page page=null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (System.ArgumentException)
            {
                caught = true;
            }
            
            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }
            
            if (page!=null)
            {
                page.Delete(0);
            }
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_VerifyThatUnknownStencilsAreDetected()
        {
            string non_existent_stencil = "foobar.vss";

            var doc = this.GetNewDoc();
            var page_node = new Page();
            var master_node_0 = page_node.Shapes.Drop(this.rectangle, non_existent_stencil, 3, 3);

            IVisio.Page page = null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (AutomationException)
            {
                caught = true;
            }
            
            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }

            if (page!=null)
            {
                page.Delete(0);                
            }
            doc.Close(true);
        }

        [TestMethod]
        public void Dom_DrawAndDrop()
        {
            var doc = this.GetNewDoc();
            var page_node = new Page();

            var rect0 = new VisioAutomation.Drawing.Rectangle(3, 4, 7, 8);
            var rect1 = new VisioAutomation.Drawing.Rectangle(8, 1, 9, 5);

            // Draw and Drop two rectangles in the same place
            var s0 = page_node.Shapes.Drop(this.rectangle, this.basic_u_vss, rect0);
            var s1 = page_node.Shapes.DrawRectangle(rect0);

            // Draw and Drop two rectangles in the same place
            var s2 = page_node.Shapes.Drop(this.rectangle, this.basic_u_vss, rect1);
            var s3 = page_node.Shapes.DrawRectangle(rect1);

            // Render the page
            var page = page_node.Render(doc);

            // Verify the locations and sizes
            var shapeids = new int[] {
                s0.VisioShapeID, 
                s1.VisioShapeID, 
                s2.VisioShapeID, 
                s3.VisioShapeID };

            var xfrms = VA.Shapes.ShapeXFormCells.GetCells(page, shapeids);

            Assert.AreEqual(xfrms[1].PinX, xfrms[0].PinX);
            Assert.AreEqual(xfrms[1].PinY, xfrms[0].PinY);

            Assert.AreEqual(xfrms[1].Width, xfrms[0].Width);
            Assert.AreEqual(xfrms[1].Height, xfrms[0].Height);

            Assert.AreEqual(xfrms[3].PinX,   xfrms[2].PinX);
            Assert.AreEqual(xfrms[3].PinY,   xfrms[2].PinY);
            Assert.AreEqual(xfrms[3].Width,  xfrms[2].Width);
            Assert.AreEqual(xfrms[3].Height, xfrms[2].Height);

            doc.Close(true);
        }
    }
}
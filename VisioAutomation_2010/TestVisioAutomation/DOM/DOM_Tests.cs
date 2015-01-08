using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    public static class CommonStencils
    {
        public class BaseStencil
        {
            public string Name;

            public BaseStencil(string n)
            {
                this.Name = n;
            }
        }

        public class BasicDef : BaseStencil
        {
            public BasicDef() : base("basic_u.vss")
            {
            }

            public string Rectangle = "Rectangle";
        }

        public class ConnectorsDef : BaseStencil
        {
            public ConnectorsDef()
                : base("connec_u.vss")
            {
            }

            public string Dynamic_Connector = "Dynamic Connector";
        }

        public class OrgChartDef : BaseStencil
        {
            public OrgChartDef()
                : base("orgch_u.vst")
            {
            }
            public string Position = "Position";
        }

        public class OrgChartBeltDef : BaseStencil
        {
            public OrgChartBeltDef()
                : base("orgch_u.vst")
            {
            }
            public string Position_Belt = "Position Belt";
        }

        public static BasicDef Basic = new BasicDef();
        public static ConnectorsDef Connectors = new ConnectorsDef();
        public static OrgChartDef OrgChart = new OrgChartDef();
        public static OrgChartBeltDef OrgChartBelt = new OrgChartBeltDef();
    }


    [TestClass]
    public class DOM_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void DOM_EmptyRendering()
        {
            // Rendering a DOM should not change the page count
            // Empty DOMs do not add any shapes
            var app = this.GetVisioApplication();

            var page_node = new VA.DOM.Page();
            var doc = this.GetNewDoc();
            page_node.Render(app.ActiveDocument);
            Assert.AreEqual(0, app.ActivePage.Shapes.Count);
            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_RenderPageToDocument()
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

        [TestMethod]
        public void DOM_RenderDocumentToApplication()
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
        public void DOM_DrawSimpleShape()
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

        [TestMethod]
        public void DOM_DropShapes()
        {
            // Render it
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var stencil = app.Documents.OpenStencil(CommonStencils.Basic.Name);
            var rectmaster = stencil.Masters[CommonStencils.Basic.Rectangle];


            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();

            shape_nodes.DrawRectangle(0, 0, 1, 1);
            shape_nodes.Drop(rectmaster, 3, 3);

            shape_nodes.Render(app.ActivePage);

            app.ActiveDocument.Close(true);
        }

        [TestMethod]
        public void DOM_CustomProperties()
        {
            // Create the doc
            var shape_nodes = new VA.DOM.ShapeList();
            var vrect1 = new VA.DOM.Rectangle(1, 1, 9, 9);
            vrect1.Text = new VA.Text.Markup.TextElement("HELLO WORLD");

            vrect1.CustomProperties = new Dictionary<string, VA.Shapes.CustomProperties.CustomPropertyCells>();

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
        public void DOM_DrawOrgChart()
        {
            var app = this.GetVisioApplication();
            var vis_ver = VA.Application.ApplicationHelper.GetApplicationVersion(app);

            // How to draw using a Template instead of a doc and a stencil
            string orgchart_vst = vis_ver.Major >= 15 ? CommonStencils.OrgChartBelt.Name : CommonStencils.OrgChart.Name;
            string position_master_name = vis_ver.Major >= 15 ? CommonStencils.OrgChartBelt.Position_Belt : CommonStencils.OrgChart.Position;

            var doc_node = new VA.DOM.Document(orgchart_vst, IVisio.VisMeasurementSystem.visMSUS);
            var page_node = new VA.DOM.Page();
            doc_node.Pages.Add(page_node);

            // Have to be smart about selecting the right master with Visio 2013

            var s1 = new VisioAutomation.DOM.Shape(position_master_name, null, new VA.Drawing.Point(3, 8));
            page_node.Shapes.Add(s1);

            var s2 = new VisioAutomation.DOM.Shape(position_master_name, null, new VA.Drawing.Point(0, 4));
            page_node.Shapes.Add(s2);

            var s3 = new VisioAutomation.DOM.Shape(position_master_name, null, new VA.Drawing.Point(6, 4));
            page_node.Shapes.Add(s3);

            page_node.Shapes.Connect(CommonStencils.Connectors.Dynamic_Connector, CommonStencils.Connectors.Name, s1, s2);
            page_node.Shapes.Connect(CommonStencils.Connectors.Dynamic_Connector, CommonStencils.Connectors.Name, s1, s3);

            var doc = doc_node.Render(app);

            //doc.Close(true);
        }

        [TestMethod]
        public void DOM_DrawEmpty()
        {
            // Verify that an empty DOM page can be created and rendered
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            page_node.Size = new VA.Drawing.Size(5, 5);
            var page = page_node.Render(doc);

            Assert.AreEqual(0, page.Shapes.Count);
            Assert.AreEqual(new VA.Drawing.Size(5, 5), VisioAutomationTest.GetPageSize(page));

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DrawLine()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
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
        public void DOM_DrawBezier()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var bez_node_0 = page_node.Shapes.DrawBezier(new double[] { 1, 2, 3, 3, 6, 3, 3, 4 });

            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, bez_node_0.VisioShapeID);
            Assert.IsNotNull(bez_node_0.VisioShape);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_DropMaster()
        {

            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil(CommonStencils.Basic.Name);
            var master1 = stencil.Masters[CommonStencils.Basic.Rectangle];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, CommonStencils.Basic.Name, 5, 5);

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
        public void DOM_FormatShape()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil(CommonStencils.Basic.Name);
            var master1 = stencil.Masters[CommonStencils.Basic.Rectangle];

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
        public void DOM_ConnectShapes()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();

            var basic_stencil = doc.Application.Documents.OpenStencil(CommonStencils.Basic.Name);
            var basic_masters = basic_stencil.Masters;
            var connectors_stencil = doc.Application.Documents.OpenStencil(CommonStencils.Connectors.Name);
            var connectors_masters = connectors_stencil.Masters;

            var master1 = basic_masters[CommonStencils.Basic.Rectangle];
            var master2 = connectors_masters[CommonStencils.Connectors.Dynamic_Connector];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(master1, 6, 5);
            var dc = page_node.Shapes.Connect(master2, master_node_0, master_node_1);

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_ConnectShapes2()
        {
            // Deferred means that the stencils (and thus masters) are loaded when rendering
            // and are no loaded by the caller before Render() is called

            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, CommonStencils.Basic.Name, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, CommonStencils.Basic.Name, 6, 5);
            var dc = page_node.Shapes.Connect(CommonStencils.Connectors.Name, CommonStencils.Connectors.Name, master_node_0, master_node_1);
            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void DOM_VerifyThatUnknownMastersAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle + "_XXX", CommonStencils.Basic.Name, 3, 3);

            IVisio.Page page=null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (VA.AutomationException)
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
        public void DOM_VerifyThatUnknownStencilsAreDetected()
        {
            string non_existent_stencil = "foobar.vss";

            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, non_existent_stencil, 3, 3);

            IVisio.Page page = null;
            bool caught = false;
            try
            {
                page = page_node.Render(doc);
            }
            catch (VA.AutomationException)
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
        public void DOM_DrawAndDrop()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();

            var rect0 = new VA.Drawing.Rectangle(3, 4, 7, 8);
            var rect1 = new VA.Drawing.Rectangle(8, 1, 9, 5);
            var dropped_shape0 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, CommonStencils.Basic.Name, rect0);
            var drawn_shape0 = page_node.Shapes.DrawRectangle(rect0);

            var dropped_shape1 = page_node.Shapes.Drop(CommonStencils.Basic.Rectangle, CommonStencils.Basic.Name, rect1);
            var drawn_shape1 = page_node.Shapes.DrawRectangle(rect1);

            var page = page_node.Render(doc);

            var xfrms = VA.Shapes.XFormCells.GetCells(page,
                                                        new int[] { dropped_shape0.VisioShapeID, 
                                                            drawn_shape0.VisioShapeID, 
                                                            dropped_shape1.VisioShapeID, 
                                                            drawn_shape1.VisioShapeID });

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
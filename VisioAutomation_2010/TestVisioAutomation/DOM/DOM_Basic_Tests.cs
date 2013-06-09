using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class DOM_Basic_Tests : VisioAutomationTest
    {

        [TestMethod]
        public void DrawEmptyDOM()
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
        public void DrawDOMSimpleNonMasterShapes()
        {
            this.DrawDOMBezier();
            this.DrawDOMLine();                    
        }

        public void DrawDOMLine()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var line_node_0 = page_node.Shapes.DrawLine(1, 1, 3, 3);
            var page = page_node.Render(doc);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, line_node_0.VisioShapeID);
            Assert.IsNotNull(line_node_0.VisioShape);
            Assert.AreEqual(2.0, line_node_0.VisioShape.CellsU["PinX"].Result[IVisio.VisUnitCodes.visNoCast]);
            page.Delete(0);
            doc.Close(true);
        }

        public void DrawDOMBezier()
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
        public void DropDOMMaster()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 5, 5);

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
        public void FormatDOMShape()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

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
        public void ConnectDOMShapes()
        {
            // Deferred means that instead of passing
            // an IVisio Master object, that 
            // the name of the master and stencil are used
            this.ConnectDOMShapesNonDeferred();
            this.ConnectDOMShapesDeferred();
        }

        public void ConnectDOMShapesNonDeferred()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var basic_stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var basic_masters = basic_stencil.Masters;

            var connectors_stencil = doc.Application.Documents.OpenStencil("connec_u.vss");
            var connectors_masters = connectors_stencil.Masters;

            var master1 = basic_masters["Rectangle"];
            var master2 = connectors_masters["Dynamic Connector"];

            var master_node_0 = page_node.Shapes.Drop(master1, 3, 3);
            var master_node_1 = page_node.Shapes.Drop(master1, 6, 5);
            var dc = page_node.Shapes.Connect(master2, master_node_0, master_node_1);

            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        public void ConnectDOMShapesDeferred()
        {
            // Deferred means that the stencils (and thus masters) are loaded when rendering
            // and are no loaded by the caller before Render() is called

            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 3, 3);
            var master_node_1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", 6, 5);
            var dc = page_node.Shapes.Connect("Dynamic Connector", "connec_u.vss", master_node_0, master_node_1);
            var page = page_node.Render(doc);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
            doc.Close(true);
        }

        [TestMethod]
        public void VerifyUnknownMastersAndStencils()
        {
            this.VerifyThatUnknownMastersAreDetected();
            this.VerifyThatUnknownStencilsAreDetected();                    
        }

        public void VerifyThatUnknownMastersAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("RectangleXXX", "basic_u.vss", 3, 3);

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


        public void VerifyThatUnknownStencilsAreDetected()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();
            var master_node_0 = page_node.Shapes.Drop("Rectangle", "basic_uXXX.vss", 3, 3);

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
        public void DroppingAndDrawingInDOMWorkTogether()
        {
            var doc = this.GetNewDoc();
            var page_node = new VA.DOM.Page();

            var rect0 = new VA.Drawing.Rectangle(3, 4, 7, 8);
            var rect1 = new VA.Drawing.Rectangle(8, 1, 9, 5);
            var dropped_shape0 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", rect0);
            var drawn_shape0 = page_node.Shapes.DrawRectangle(rect0);

            var dropped_shape1 = page_node.Shapes.Drop("Rectangle", "basic_u.vss", rect1);
            var drawn_shape1 = page_node.Shapes.DrawRectangle(rect1);

            var page = page_node.Render(doc);

            var xfrms = VA.Layout.XFormCells.GetCells(page,
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
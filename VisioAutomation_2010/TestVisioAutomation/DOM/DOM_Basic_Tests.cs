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
            var page = this.GetNewPage();

            var dom = new VA.DOM.ShapeCollection();
            page.SetSize(new VA.Drawing.Size(5, 5));
            dom.Render(page);

            Assert.AreEqual(0, page.Shapes.Count);
            Assert.AreEqual(new VA.Drawing.Size(5, 5), page.GetSize());

            page.Delete(0);
        }

        [TestMethod]
        public void DropLine()
        {
            var page = this.GetNewPage();

            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_line_0 = domshapescol.DrawLine(1, 1, 3, 3);
            domshapescol.Render(page);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, dom_line_0.VisioShapeID);
            Assert.IsNotNull(dom_line_0.VisioShape);
            Assert.AreEqual(2.0, dom_line_0.VisioShape.CellsU["PinX"].Result[IVisio.VisUnitCodes.visNoCast]);
            page.Delete(0);
        }

        [TestMethod]
        public void DropBezier()
        {
            var page = this.GetNewPage();

            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_bez_0 = domshapescol.DrawBezier(new double[] {1, 2, 3, 3, 6, 3, 3, 4});

            domshapescol.Render(page);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, dom_bez_0.VisioShapeID);
            Assert.IsNotNull(dom_bez_0.VisioShape);
            page.Delete(0);
        }

        [TestMethod]
        public void DropMaster()
        {
            var page = this.GetNewPage();
            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop(master1, 3, 3);
            var dom_master_1 = domshapescol.Drop("Rectangle", "basic_u.vss", 5, 5);

            domshapescol.Render(page);

            Assert.AreEqual(2, page.Shapes.Count);

            // Verify that the shapes created both have IDs and shape objects associated with them
            Assert.AreNotEqual(0, dom_master_0.VisioShapeID);
            Assert.AreNotEqual(0, dom_master_1.VisioShapeID);
            Assert.IsNotNull(dom_master_0.VisioShape);
            Assert.IsNotNull(dom_master_1.VisioShape);
            page.Delete(0);
        }

        [TestMethod]
        public void ShapeFormat()
        {
            var page = this.GetNewPage();
            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop(master1, 3, 3);
            var dom_bez_0 = domshapescol.DrawBezier(new double[] {1, 2, 3, 3, 6, 3, 3, 4});
            var dom_line_0 = domshapescol.DrawLine(1, 1, 3, 3);

            dom_master_0.Cells.LineWeight = 0.1;
            dom_bez_0.Cells.LineWeight = 0.3;
            dom_line_0.Cells.LineWeight = 0.5;

            domshapescol.Render(page);

            Assert.AreEqual(3, page.Shapes.Count);
            page.Delete(0);
        }

        [TestMethod]
        public void Connect()
        {
            var page = this.GetNewPage();
            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];
            var master2 = stencil.Masters["Dynamic Connector"];

            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop(master1, 3, 3);
            var dom_master_1 = domshapescol.Drop(master1, 6, 5);
            var dc = domshapescol.Connect(master2, dom_master_0, dom_master_1);

            domshapescol.Render(page);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
        }

        [TestMethod]
        public void ConnectDeferred()
        {
            var page = this.GetNewPage();
            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop("Rectangle", "basic_u.vss", 3, 3);
            var dom_master_1 = domshapescol.Drop("Rectangle", "basic_u.vss", 6, 5);
            var dc = domshapescol.Connect("Dynamic Connector", "basic_u.vss", dom_master_0, dom_master_1);

            domshapescol.Render(page);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
        }


        [TestMethod]
        public void DropUnknownMaster()
        {
            var page = this.GetNewPage();
            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop("RectangleXXX", "basic_u.vss", 3, 3);

            bool caught = false;
            try
            {
                domshapescol.Render(page);
            }
            catch (VA.AutomationException)
            {
                caught = true;
            }


            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }


            page.Delete(0);
        }


        [TestMethod]
        public void DropUnknownStencil()
        {
            var page = this.GetNewPage();
            var domshapescol = new VA.DOM.ShapeCollection();
            var dom_master_0 = domshapescol.Drop("RectangleXXX", "basic_uXXX.vss", 3, 3);

            bool caught = false;
            try
            {
                domshapescol.Render(page);
            }
            catch (VA.AutomationException)
            {
                caught = true;
            }


            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }


            page.Delete(0);
        }

        [TestMethod]
        public void VerifyDropVersusDraw()
        {
            var page = this.GetNewPage();
            var dom = new VA.DOM.ShapeCollection();
            var rect = new VA.Drawing.Rectangle(3, 4, 7, 8);
            var dropped_shape = dom.Drop("Rectangle", "basic_u.vss", rect);
            var drawn_shape = dom.DrawRectangle(rect);
            dom.Render(page);

            var xfrms = VA.Layout.LayoutHelper.GetXForm(page,
                                                        new int[] {dropped_shape.VisioShapeID, drawn_shape.VisioShapeID});

            Assert.AreEqual(xfrms[1].PinX,xfrms[0].PinX);
            Assert.AreEqual(xfrms[1].PinY, xfrms[0].PinY);

            Assert.AreEqual(xfrms[1].Width, xfrms[0].Width);
            Assert.AreEqual(xfrms[1].Height, xfrms[0].Height);

            page.Delete(0);
        }

        [TestMethod]
        public void VerifyDropVersusDraw2()
        {
            var page = this.GetNewPage();
            var domshapescol = new VA.DOM.ShapeCollection();
            var rect0 = new VA.Drawing.Rectangle(3, 4, 7, 8);
            var rect1 = new VA.Drawing.Rectangle(8, 1, 9, 5);
            var dropped_shape0 = domshapescol.Drop("Rectangle", "basic_u.vss", rect0);
            var drawn_shape0 = domshapescol.DrawRectangle(rect0);

            var dropped_shape1 = domshapescol.Drop("Rectangle", "basic_u.vss", rect1);
            var drawn_shape1 = domshapescol.DrawRectangle(rect1);

            domshapescol.Render(page);

            var xfrms = VA.Layout.LayoutHelper.GetXForm(page,
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

            page.Delete(0);
        }
    }
}
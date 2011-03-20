using System;
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

            var dom = new VA.DOM.Document();
            dom.PageSettings.Size = new VA.Drawing.Size(5, 5);
            dom.Render(page);

            Assert.AreEqual(0, page.Shapes.Count);
            Assert.AreEqual(new VA.Drawing.Size(5, 5), page.GetSize());

            page.Delete(0);
        }

        [TestMethod]
        public void DropLine()
        {
            var page = this.GetNewPage();

            IVisio.Shape s1;
            var dom = new VA.DOM.Document();
            var dom_line_0 = dom.DrawLine(1, 1, 3, 3);
            dom.Render(page);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, dom_line_0.ShapeID);
            Assert.IsNotNull(dom_line_0.VisioShape);
            Assert.AreEqual(2.0, dom_line_0.VisioShape.CellsU["PinX"].Result[IVisio.VisUnitCodes.visNoCast]);
            page.Delete(0);
        }

        [TestMethod]
        public void DropBezier()
        {
            var page = this.GetNewPage();

            IVisio.Shape s1;
            var dom = new VA.DOM.Document();
            var dom_bez_0 = dom.DrawBezier(new double[] {1, 2, 3, 3, 6, 3, 3, 4});

            dom.Render(page);

            Assert.AreEqual(1, page.Shapes.Count);
            Assert.AreNotEqual(0, dom_bez_0.ShapeID);
            Assert.IsNotNull(dom_bez_0.VisioShape);
            page.Delete(0);
        }

        [TestMethod]
        public void DropMaster()
        {
            var page = this.GetNewPage();
            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            IVisio.Shape s1;
            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop(master1, 3, 3);
            var dom_master_1 = dom.Drop("Rectangle", "basic_u.vss", 5, 5);

            dom.Render(page);

            Assert.AreEqual(2, page.Shapes.Count);

            // Unless overriden, Dropped masters don't have their shapes collected, only their shapeids
            Assert.AreNotEqual(0, dom_master_0.ShapeID);
            Assert.AreNotEqual(0, dom_master_1.ShapeID);
            Assert.IsNull(dom_master_0.VisioShape);
            Assert.IsNull(dom_master_1.VisioShape);
            page.Delete(0);
        }

        [TestMethod]
        public void ShapeFormat()
        {
            var page = this.GetNewPage();
            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var master1 = stencil.Masters["Rectangle"];

            IVisio.Shape s1;
            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop(master1, 3, 3);
            var dom_bez_0 = dom.DrawBezier(new double[] {1, 2, 3, 3, 6, 3, 3, 4});
            var dom_line_0 = dom.DrawLine(1, 1, 3, 3);

            dom_master_0.ShapeCells.LineWeight = 0.1;
            dom_bez_0.ShapeCells.LineWeight = 0.3;
            dom_line_0.ShapeCells.LineWeight = 0.5;

            dom.Render(page);

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

            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop(master1, 3, 3);
            var dom_master_1 = dom.Drop(master1, 6, 5);
            var dc = dom.Connect(master2, dom_master_0, dom_master_1);

            dom.Render(page);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
        }

        [TestMethod]
        public void ConnectDeferred()
        {
            var page = this.GetNewPage();
            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop("Rectangle", "basic_u.vss", 3, 3);
            var dom_master_1 = dom.Drop("Rectangle", "basic_u.vss", 6, 5);
            var dc = dom.Connect("Dynamic Connector", "basic_u.vss", dom_master_0, dom_master_1);

            dom.Render(page);

            Assert.AreEqual(3, page.Shapes.Count);

            page.Delete(0);
        }


        [TestMethod]
        public void DropUnknownMaster()
        {
            var page = this.GetNewPage();
            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop("RectangleXXX", "basic_u.vss", 3, 3);

            bool caught = false;
            try
            {
                dom.Render(page);
            }
            catch (VA.AutomationException exc)
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
            var dom = new VA.DOM.Document();
            var dom_master_0 = dom.Drop("RectangleXXX", "basic_uXXX.vss", 3, 3);

            bool caught = false;
            try
            {
                dom.Render(page);
            }
            catch (VA.AutomationException exc)
            {
                caught = true;
            }


            if (caught == false)
            {
                Assert.Fail("Expected an AutomationException");
            }


            page.Delete(0);
        }

    }
}
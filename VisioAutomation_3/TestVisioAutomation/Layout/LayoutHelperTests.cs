using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class LayoutHelper : VisioAutomationTest
    {
        [TestMethod]
        public void DistributeX()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            var s1 = page.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = page.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = page.DrawRectangle(4.5, 2.5, 6, 3.5);

            var shapeids = new int[] { s1.ID, s2.ID, s3.ID };

            VA.Layout.LayoutHelper.DistributeWithSpacing(page, shapeids, VA.Drawing.Axis.XAxis, 1.0);

            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin);
            Assert.AreEqual(new VA.Drawing.Point(2.5, 3.25), xforms[1].Pin);
            Assert.AreEqual(new VA.Drawing.Point(4.5, 3), xforms[2].Pin);

            doc.Close(true);
        }

        [TestMethod]
        public void DistributeY()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            var s1 = page.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = page.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = page.DrawRectangle(4.5, 2.5, 6, 3.5);

            var shapeids = new int[] { s1.ID, s2.ID, s3.ID };

            VA.Layout.LayoutHelper.DistributeWithSpacing(page, shapeids, VA.Drawing.Axis.YAxis, 1.0);

            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin);
            Assert.AreEqual(new VA.Drawing.Point(2.25, 4.75), xforms[1].Pin);
            Assert.AreEqual(new VA.Drawing.Point(5.25, 3), xforms[2].Pin);

            doc.Close(true);
        }


        [TestMethod]
        public void AlignV()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            var s1 = page.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = page.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = page.DrawRectangle(4.5, 2.5, 6, 3.5);

            var shapeids = new int[] { s1.ID, s2.ID, s3.ID };

            VA.Layout.LayoutHelper.AlignTo(page, shapeids, VA.Drawing.AlignmentVertical.Center, 1.5);

            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.5), xforms[0].Pin);
            Assert.AreEqual(new VA.Drawing.Point(2.25, 1.5), xforms[1].Pin);
            Assert.AreEqual(new VA.Drawing.Point(5.25, 1.5), xforms[2].Pin);

            doc.Close(true);
        }


        [TestMethod]
        public void Scripting_AlignH()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            var s1 = page.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = page.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = page.DrawRectangle(4.5, 2.5, 6, 3.5);

            var shapeids = new int[] { s1.ID, s2.ID, s3.ID };

            VA.Layout.LayoutHelper.AlignTo(page, shapeids, VA.Drawing.AlignmentHorizontal.Center, 0.10);

            var xforms = VA.Layout.LayoutHelper.GetXForm(page, shapeids);
            Assert.AreEqual(new VA.Drawing.Point(0.1, 1.25), xforms[0].Pin);
            Assert.AreEqual(new VA.Drawing.Point(0.1, 3.25), xforms[1].Pin);
            Assert.AreEqual(new VA.Drawing.Point(0.1, 3.0), xforms[2].Pin);
            
            doc.Close(true);
        }

    }
}
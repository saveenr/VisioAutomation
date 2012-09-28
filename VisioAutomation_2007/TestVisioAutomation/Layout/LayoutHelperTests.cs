using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    public static class HExtensins
    {
        public static VA.Drawing.Point Pin( this VA.Layout.XFormCells xthis)
        {
            return new VA.Drawing.Point(xthis.PinX.Result, xthis.PinY.Result);
        }
  
    }

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
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(2.5, 3.25), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(4.5, 3), xforms[2].Pin());

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
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(2.25, 4.75), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(5.25, 3), xforms[2].Pin());

            doc.Close(true);
        }

        [TestMethod]
        public void Sort1()
        {
            var app = this.GetVisioApplication();
            var doc = this.GetNewDoc();
            var page = app.ActivePage;

            var s1 = page.DrawRectangle(2, 2, 3, 3);
            var s2 = page.DrawRectangle(1, 1, 2, 2);
            var s3 = page.DrawRectangle(4, 4, 4, 4);
            var s4 = page.DrawRectangle(3, 3, 3, 3);

            s1.Text = "A";
            s2.Text = "B";
            s3.Text = "C";
            s4.Text = "D";

            var shapes = new[] {s1, s2, s3, s4};
            var shapeids = shapes.Select(s=>s.ID).ToList();

            var sorted_shapeids = VA.Layout.LayoutHelper.SortShapesByPosition(page, shapeids, VA.Layout.XFormPosition.PinX);

            var sorted_shapes = sorted_shapeids.Select(id => page.Shapes.get_ItemFromID(id)).ToList();
            var text = string.Join("", sorted_shapes.Select(s => s.Text));
            Assert.AreEqual("BADC",text);

            sorted_shapeids = VA.Layout.LayoutHelper.SortShapesByPosition(page, shapeids, VA.Layout.XFormPosition.PinY);
            sorted_shapes = sorted_shapeids.Select(id => page.Shapes.get_ItemFromID(id)).ToList();
            text = string.Join("", sorted_shapes.Select(s => s.Text));
            Assert.AreEqual("BADC",text);

            doc.Close(true);
        }

    }
}
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using System.Linq;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeGeometryHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Test_AddGeometrySection()
        {
            var page = this.GetNewPage();
            var shape = page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1,shape.GeometryCount);

            var geom1 = new VA.ShapeGeometry.GeometrySection();
            geom1.NoFill = "true";
            geom1.AddMoveTo("-1", "-1");
            geom1.AddLineTo("1", "0");
            geom1.AddLineTo("1", "1");
            geom1.AddLineTo("0", "1");
            geom1.AddLineTo("0", "0");

            geom1.Render(shape);

            page.Delete(0);
        }

        [TestMethod]
        public void Test_RemoveGeometry()
        {
            var page = this.GetNewPage();

            // create simple rectangle
            // this should always have a single geometry row
            var shape1 = page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1, shape1.GeometryCount);

            // delete that single geometry row
            VA.ShapeGeometry.ShapeGeometryHelper.DeleteGeometrySection(shape1,0);
            Assert.AreEqual(0, shape1.GeometryCount);

            // now create a shape with two geometry rows
            var shape2 = page.DrawRectangle(4, 4, 5, 5);
            Assert.AreEqual(1, shape2.GeometryCount);

            var geom1 = new VA.ShapeGeometry.GeometrySection();
            geom1.NoFill = "true";
            geom1.AddMoveTo("-1", "-1");
            geom1.AddLineTo("1", "0");
            geom1.AddLineTo("1", "1");
            geom1.AddLineTo("0", "1");
            geom1.AddLineTo("0", "0");
            geom1.Render(shape2);
            Assert.AreEqual(2, shape2.GeometryCount);

            // remove all the geometry rows
            VA.ShapeGeometry.ShapeGeometryHelper.DeleteGeometry(shape2);
            Assert.AreEqual(0, shape2.GeometryCount);

            page.Delete(0);
        }

    }
}
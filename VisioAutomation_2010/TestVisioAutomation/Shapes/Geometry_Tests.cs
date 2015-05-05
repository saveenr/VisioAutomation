using Microsoft.VisualStudio.TestTools.UnitTesting;
using VAGEOM = VisioAutomation.Shapes.Geometry;

namespace TestVisioAutomation.Shapes
{
    [TestClass]
    public class Geometry_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Geometry_AddGeometrySection()
        {
            var page = this.GetNewPage();
            var shape = page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1,shape.GeometryCount);

            var geom1 = new VAGEOM.GeometrySection();
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
        public void Geometry_DeleteGeometry()
        {
            var page = this.GetNewPage();

            // create a shape with two geometry rows
            var shape2 = page.DrawRectangle(4, 4, 5, 5);
            Assert.AreEqual(1, shape2.GeometryCount);

            var geom1 = new VAGEOM.GeometrySection();
            geom1.NoFill = "true";
            geom1.AddMoveTo("-1", "-1");
            geom1.AddLineTo("1", "0");
            geom1.AddLineTo("1", "1");
            geom1.AddLineTo("0", "1");
            geom1.AddLineTo("0", "0");
            geom1.Render(shape2);
            Assert.AreEqual(2, shape2.GeometryCount);

            // remove all the geometry
            VAGEOM.GeometryHelper.Delete(shape2);
            Assert.AreEqual(0, shape2.GeometryCount);

            page.Delete(0);
        }
    }
}
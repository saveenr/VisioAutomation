using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class GeometryTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Geometry_AddGeometrySection()
        {
            var page = this.GetNewPage();
            var shape = page.DrawRectangle(1, 1, 3, 3);
            MUT.Assert.AreEqual(1,shape.GeometryCount);

            var geom1 = new GeometrySection();
            geom1.NoFill = "true";
            geom1.AddMoveTo("-1", "-1");
            geom1.AddLineTo("1", "0");
            geom1.AddLineTo("1", "1");
            geom1.AddLineTo("0", "1");
            geom1.AddLineTo("0", "0");

            short sec_index = geom1.Render(shape);

            // Render must return the section index Visio assigned to the newly-added
            // geometry section (issue #128). The rectangle started with one geometry
            // section, so Render adds a second at visSectionFirstComponent + 1.
            MUT.Assert.AreEqual(2, shape.GeometryCount);
            MUT.Assert.AreEqual((short)((short)IVisio.VisSectionIndices.visSectionFirstComponent + 1), sec_index);

            page.Delete(0);
        }

        [MUT.TestMethod]
        public void Geometry_DeleteGeometry()
        {
            var page = this.GetNewPage();

            // create a shape with two geometry rows
            var shape2 = page.DrawRectangle(4, 4, 5, 5);
            MUT.Assert.AreEqual(1, shape2.GeometryCount);

            var geom1 = new GeometrySection();
            geom1.NoFill = "true";
            geom1.AddMoveTo("-1", "-1");
            geom1.AddLineTo("1", "0");
            geom1.AddLineTo("1", "1");
            geom1.AddLineTo("0", "1");
            geom1.AddLineTo("0", "0");
            geom1.Render(shape2);
            MUT.Assert.AreEqual(2, shape2.GeometryCount);

            // remove all the geometry
            GeometryHelper.Delete(shape2);
            MUT.Assert.AreEqual(0, shape2.GeometryCount);

            page.Delete(0);
        }
    }
}
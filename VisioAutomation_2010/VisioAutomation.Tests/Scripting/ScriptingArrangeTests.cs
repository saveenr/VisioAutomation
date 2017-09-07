using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingArrangeTests : VisioAutomationTest
    {     
        [TestMethod]
        public void Scripting_Distribute()
        {
            var client = this.GetScriptingClient();
            var pagesize = new VA.Geometry.Size(4, 4);

            client.Document.New();
            client.Page.New(pagesize, false);

            var size1 = new VA.Geometry.Size(0.5, 0.5);
            var size2 = new VA.Geometry.Size(1.0, 1.0);
            var size3 = new VA.Geometry.Size(1.5, 1.5);

            var r1 = new VA.Geometry.Rectangle(new VA.Geometry.Point(1, 1), size1);
            var r2 = new VA.Geometry.Rectangle(new VA.Geometry.Point(2, 2), size2);
            var r3 = new VA.Geometry.Rectangle(new VA.Geometry.Point(4, 4), size3);

            var s1 = client.Draw.Rectangle(r1);
            var s2 = client.Draw.Rectangle(r2);
            var s3 = client.Draw.Rectangle(r3);

            client.Selection.SelectNone();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioScripting.Models.TargetShapes();

            client.Distribute.DistributeHorizontal(targets, VisioScripting.Models.AlignmentHorizontal.Center);

            var shapeids = new[] {s1.ID, s2.ID, s3.ID };
            VisioAutomation.Shapes.ShapeXFormCells.GetFormulas(client.Page.Get(),shapeids);

            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Distribute_With_Spacing()
        {
            var client = this.GetScriptingClient();
            var pagesize = new VA.Geometry.Size(4, 4);

            client.Document.New();
            client.Page.New(pagesize, false);

            var size1 = new VA.Geometry.Size(0.5, 0.5);
            var size2 = new VA.Geometry.Size(1.0, 1.0);
            var size3 = new VA.Geometry.Size(1.5, 1.5);

            var r1 = new VA.Geometry.Rectangle(new VA.Geometry.Point(1, 1), size1);
            var r2 = new VA.Geometry.Rectangle(new VA.Geometry.Point(2, 2), size2);
            var r3 = new VA.Geometry.Rectangle(new VA.Geometry.Point(4, 4), size3);

            var s1 = client.Draw.Rectangle(r1);
            var s2 = client.Draw.Rectangle(r2);
            var s3 = client.Draw.Rectangle(r3);

            client.Selection.SelectNone();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioScripting.Models.TargetShapes();
            client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.XAxis , 0.25);
            client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.YAxis, 1.0);

            var shapeids = new[] { s1.ID, s2.ID, s3.ID };
            var out_xfrms = VisioAutomation.Shapes.ShapeXFormCells.GetResults(client.Page.Get(), shapeids);
            var out_positions = out_xfrms.Select(xfrm => TestExtensions.ToPoint(xfrm.PinX.Value, xfrm.PinY.Value)).ToArray();

            Assert.AreEqual(1.25, out_positions[0].X);
            Assert.AreEqual(1.25, out_positions[0].Y);
            Assert.AreEqual(2.25, out_positions[1].X);
            Assert.AreEqual(3.00, out_positions[1].Y);
            Assert.AreEqual(3.75, out_positions[2].X);
            Assert.AreEqual(5.25, out_positions[2].Y);

            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Nudge2()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Geometry.Size(4, 4), false);

            var size1 = new VA.Geometry.Size(0.5, 0.5);
            var size2 = new VA.Geometry.Size(1.0, 1.0);
            var size3 = new VA.Geometry.Size(1.5, 1.5);

            var r1 = new VA.Geometry.Rectangle(new VA.Geometry.Point(1, 1), size1);
            var r2 = new VA.Geometry.Rectangle(new VA.Geometry.Point(2, 2), size2);
            var r3 = new VA.Geometry.Rectangle(new VA.Geometry.Point(4, 4), size3);

            var s1 = client.Draw.Rectangle(r1);
            var s2 = client.Draw.Rectangle(r2);
            var s3 = client.Draw.Rectangle(r3);

            client.Selection.SelectNone();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioScripting.Models.TargetShapes();

            client.Arrange.Nudge(targets, 0.50, -0.25);

            var shapeids = new[] { s1.ID, s2.ID, s3.ID };
            var xforms = VisioAutomation.Shapes.ShapeXFormCells.GetResults(client.Page.Get(), shapeids);

            AssertUtil.AreEqual( (1.75, 1), xforms[0].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual( (3, 2.25), xforms[1].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual( (5.25, 4.5), xforms[2].GetPinPosResult(), 0.00001);
            client.Document.Close(true);
        }
    }
}
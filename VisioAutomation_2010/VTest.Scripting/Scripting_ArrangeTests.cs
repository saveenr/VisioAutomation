using System.Linq;
using VTest.Framework;
using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class Scripting_ArrangeTests : Framework.VTest
    {     
        [MUT.TestMethod]
        public void Scripting_Distribute()
        {
            var client = this.GetScriptingClient();
            var pagesize = new VA.Core.Size(4, 4);

            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var size1 = new VA.Core.Size(0.5, 0.5);
            var size2 = new VA.Core.Size(1.0, 1.0);
            var size3 = new VA.Core.Size(1.5, 1.5);

            var r1 = new VA.Core.Rectangle(new VA.Core.Point(1, 1), size1);
            var r2 = new VA.Core.Rectangle(new VA.Core.Point(2, 2), size2);
            var r3 = new VA.Core.Rectangle(new VA.Core.Point(4, 4), size3);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r1);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r2);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r3);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            client.Arrange.DistributeHorizontal(VisioScripting.TargetSelection.Auto, VisioScripting.Models.AlignmentHorizontal.Center);

            var shapes = new[] { s1, s2, s3 };
            var shapeids = shapes.Select(s => (int)s.ID16).ToList();
            VisioAutomation.Shapes.ShapeXFormCells.GetCells(client.Page.GetActivePage(),shapeids, VA.Core.CellValueType.Formula);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Nudge2()
        {
            var client = this.GetScriptingClient();

            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 4), false);

            var size1 = new VA.Core.Size(0.5, 0.5);
            var size2 = new VA.Core.Size(1.0, 1.0);
            var size3 = new VA.Core.Size(1.5, 1.5);

            var r1 = new VA.Core.Rectangle(new VA.Core.Point(1, 1), size1);
            var r2 = new VA.Core.Rectangle(new VA.Core.Point(2, 2), size2);
            var r3 = new VA.Core.Rectangle(new VA.Core.Point(4, 4), size3);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r1);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r2);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, r3);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            client.Arrange.Nudge(VisioScripting.TargetSelection.Auto, 0.50, -0.25);

            var shapes = new[] { s1, s2, s3 };
            var shapeids = shapes.Select(s => (int) s.ID16).ToList();
            var xforms = VisioAutomation.Shapes.ShapeXFormCells.GetCells(client.Page.GetActivePage(), shapeids, VA.Core.CellValueType.Result);

            AssertUtil.AreEqual( (1.75, 1), xforms[0].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual( (3, 2.25), xforms[1].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual( (5.25, 4.5), xforms[2].GetPinPosResult(), 0.00001);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
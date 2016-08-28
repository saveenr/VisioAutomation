using Microsoft.Office.Interop.Visio;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Drawing.Layout;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingArrangeTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Arrangement_Scenarios()
        {
            this.Scripting_Distribute();
            this.Scripting_Nudge();
        }

        public void Scripting_Distribute()
        {
            var client = this.GetScriptingClient();

            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioAutomation.Scripting.TargetShapes();

            client.Distribute.DistributeHorizontal(targets,AlignmentHorizontal.Center);

            VisioAutomation.Shapes.XFormCells.GetCells(client.Page.Get(),new[] {s1.ID, s2.ID, s3.ID });

            client.Document.Close(true);
        }

        public void Scripting_Nudge()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var targets = new VisioAutomation.Scripting.TargetShapes();

            client.Arrange.Nudge(targets,1, -1);

            var xforms = VisioAutomation.Shapes.XFormCells.GetCells(client.Page.Get(), new[] { s1.ID, s2.ID, s3.ID });

            AssertUtil.AreEqual(2.125, 0.25, xforms[0].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual(3.25, 2.25, xforms[1].GetPinPosResult(), 0.00001);
            AssertUtil.AreEqual(6.25, 2, xforms[2].GetPinPosResult(), 0.00001);
            client.Document.Close(true);
        }
    }
}
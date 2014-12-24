using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
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
            var client = GetScriptingClient();

            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            client.Arrange.Distribute(null,VA.Drawing.AlignmentHorizontal.Center);

            var xforms = client.Arrange.GetXForm(null);
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(3.1875, 3.25), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(5.25, 3), xforms[2].Pin());

            client.Document.Close(true);
        }

        public void Scripting_Nudge()
        {
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            client.Arrange.Nudge(null,1, -1);

            var xforms = client.Arrange.GetXForm(null);
            Assert.AreEqual(new VA.Drawing.Point(2.125, 0.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(3.25, 2.25), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(6.25, 2), xforms[2].Pin());
            client.Document.Close(true);
        }
    }
}
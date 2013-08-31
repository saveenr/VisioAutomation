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
            var ss = GetScriptingSession();

            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = ss.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = ss.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = ss.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.Select(s1);
            ss.Selection.Select(s2);
            ss.Selection.Select(s3);

            ss.Layout.Distribute(null,VA.Drawing.AlignmentHorizontal.Center);

            var xforms = ss.Layout.GetXForm(null);
            Assert.AreEqual(new VA.Drawing.Point(1.125, 1.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(3.1875, 3.25), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(5.25, 3), xforms[2].Pin());

            ss.Document.Close(true);
        }

        public void Scripting_Nudge()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var s1 = ss.Draw.Rectangle(1, 1, 1.25, 1.5);
            var s2 = ss.Draw.Rectangle(2, 3, 2.5, 3.5);
            var s3 = ss.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.Select(s1);
            ss.Selection.Select(s2);
            ss.Selection.Select(s3);

            ss.Layout.Nudge(null,1, -1);

            var xforms = ss.Layout.GetXForm(null);
            Assert.AreEqual(new VA.Drawing.Point(2.125, 0.25), xforms[0].Pin());
            Assert.AreEqual(new VA.Drawing.Point(3.25, 2.25), xforms[1].Pin());
            Assert.AreEqual(new VA.Drawing.Point(6.25, 2), xforms[2].Pin());
            ss.Document.Close(true);
        }
    }
}
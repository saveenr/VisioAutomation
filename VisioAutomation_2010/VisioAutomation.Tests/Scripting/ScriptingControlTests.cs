using Microsoft.VisualStudio.TestTools.UnitTesting;
using VACONTROL = VisioAutomation.Shapes.Controls;
using VAHYPERLINK = VisioAutomation.Shapes.Hyperlinks;

namespace TestVisioAutomation.Scripting
{
    [TestClass]
    public class ScriptingControlTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Controls_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VisioAutomation.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.5, 1.5);

            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.Rectangle(1.5, 3.5, 2, 4.0);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var controls0 = client.Control.Get(null);
            int found_controls = controls0.Count;
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls0[s1].Count);
            Assert.AreEqual(0, controls0[s2].Count);
            Assert.AreEqual(0, controls0[s3].Count);

            var ctrl = new VACONTROL.ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "0";
            client.Control.Add(null, ctrl);

            var controls1 = client.Control.Get(null);
            Assert.AreEqual(3, controls1.Count);
            Assert.AreEqual(1, controls1[s1].Count);
            Assert.AreEqual(1, controls1[s2].Count);
            Assert.AreEqual(1, controls1[s3].Count);

            client.Control.Delete(null, 0);
            var controls2 = client.Control.Get(null);
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls2[s1].Count);
            Assert.AreEqual(0, controls2[s2].Count);
            Assert.AreEqual(0, controls2[s3].Count);

            client.Document.Close(true);
        }
    }

    [TestClass]
    public class ScriptingHyperlinkTests : VisioAutomationTest
    {


        [TestMethod]
        public void Scripting_Hyperlinks_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VisioAutomation.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.5, 1.5);

            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.Rectangle(1.5, 3.5, 2, 4.0);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var hyperlinks0 = client.Hyperlink.Get(null);

            Assert.AreEqual(3, hyperlinks0.Count);
            Assert.AreEqual(0, hyperlinks0[s1].Count);
            Assert.AreEqual(0, hyperlinks0[s2].Count);
            Assert.AreEqual(0, hyperlinks0[s3].Count);

            var hyperlink = new VAHYPERLINK.HyperlinkCells();
            hyperlink.Address = "http://www.microsoft.com";
            client.Hyperlink.Add(null, hyperlink);

            var hyperlinks1 = client.Hyperlink.Get(null);
            Assert.AreEqual(3, hyperlinks1.Count);
            Assert.AreEqual(1, hyperlinks1[s1].Count);
            Assert.AreEqual(1, hyperlinks1[s2].Count);
            Assert.AreEqual(1, hyperlinks1[s3].Count);

            client.Hyperlink.Delete(null, 0);
            var hyperlinks2 = client.Hyperlink.Get(null);
            Assert.AreEqual(3, hyperlinks0.Count);
            Assert.AreEqual(0, hyperlinks2[s1].Count);
            Assert.AreEqual(0, hyperlinks2[s2].Count);
            Assert.AreEqual(0, hyperlinks2[s3].Count);

            client.Document.Close(true);
        }
    }
}
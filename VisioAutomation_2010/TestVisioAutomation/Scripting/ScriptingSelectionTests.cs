using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace TestVisioAutomation.Scripting
{
    [TestClass]
    public class ScriptingSelectionTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Selection_Scenarios()
        {
            var client = this.GetScriptingClient();

            var doc = client.Document.New(10, 5);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;
            var selection = active_window.Selection;
            var x1 = selection.AsEnumerable().ToDictionary(s => s);
            Assert.AreEqual(1, x1.Count);
            Assert.IsTrue(x1.ContainsKey(s4));

            client.Selection.Invert();
            var x2 = active_window.Selection.AsEnumerable().ToDictionary(s => s);
            Assert.AreEqual(3, x2.Count);
            Assert.IsTrue(x2.ContainsKey(s1));
            Assert.IsTrue(x2.ContainsKey(s2));
            Assert.IsTrue(x2.ContainsKey(s3));
            Assert.IsTrue(!x2.ContainsKey(s4));

            active_window.SelectAll();
            //app.ActiveWindows.Selection.SelectAll() selects 3 items
            var x3 = active_window.Selection.AsEnumerable().ToDictionary(s => s);
            Assert.AreEqual(4, x3.Count);

            active_window.DeselectAll();
            //app.ActiveWindows.Selection.DeselectAll() keeps all 4 selection
            var x4 = active_window.Selection.AsEnumerable().ToDictionary(s => s);
            Assert.AreEqual(0, x4.Count);

            client.Document.Close(true);
        }
    }
}
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingSelectionTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Selection_Scenarios()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(10,5);
            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;


            var selection = active_window.Selection;
            var x1 = selection.ToEnumerable().ToDictionary(s => s);
            Assert.AreEqual(1, x1.Count);
            Assert.IsTrue(x1.ContainsKey(s4));

            var targetselection = new VisioScripting.TargetActiveSelection();
            var targetwindow = new VisioScripting.TargetWindow();


            client.Selection.InvertSelection(targetwindow);

            var x2 = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            Assert.AreEqual(3, x2.Count);
            Assert.IsTrue(x2.ContainsKey(s1));
            Assert.IsTrue(x2.ContainsKey(s2));
            Assert.IsTrue(x2.ContainsKey(s3));
            Assert.IsTrue(!x2.ContainsKey(s4));

            active_window.SelectAll();
            //app.ActiveWindows.Selection.SelectAll() selects 3 items
            var x3 = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            Assert.AreEqual(4, x3.Count);

            active_window.DeselectAll();
            //app.ActiveWindows.Selection.DeselectAll() keeps all 4 selection
            var x4 = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            Assert.AreEqual(0, x4.Count);

            var targetdoc = new VisioScripting.TargetDocument();
            client.Document.CloseDocument(targetdoc, true);
        }
    }
}
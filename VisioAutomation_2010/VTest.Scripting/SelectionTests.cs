using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class SelectionTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void DrawRectangle_FollowedByActiveWindowSelection_LastDrawnShapeIsSelected()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Core.Size(10, 5);
            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            page1.DrawRectangle(0, 0, 1, 1);
            page1.DrawRectangle(1, 0, 2, 1);
            page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;
            var selected = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            MUT.Assert.AreEqual(1, selected.Count);
            MUT.Assert.IsTrue(selected.ContainsKey(s4));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void InvertSelection_OnPageWithFourShapesAndOneSelected_LeavesOtherThreeSelected()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Core.Size(10, 5);
            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;
            client.Selection.InvertSelection(VisioScripting.TargetWindow.Auto);

            var selected = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            MUT.Assert.AreEqual(3, selected.Count);
            MUT.Assert.IsTrue(selected.ContainsKey(s1));
            MUT.Assert.IsTrue(selected.ContainsKey(s2));
            MUT.Assert.IsTrue(selected.ContainsKey(s3));
            MUT.Assert.IsFalse(selected.ContainsKey(s4));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void SelectAll_OnPageWithFourShapes_SelectsAllFour()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Core.Size(10, 5);
            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            page1.DrawRectangle(0, 0, 1, 1);
            page1.DrawRectangle(1, 0, 2, 1);
            page1.DrawRectangle(0, 1, 1, 2);
            page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;
            active_window.SelectAll();

            var selected = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            MUT.Assert.AreEqual(4, selected.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DeselectAll_AfterSelectAll_LeavesNothingSelected()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Core.Size(10, 5);
            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            page1.DrawRectangle(0, 0, 1, 1);
            page1.DrawRectangle(1, 0, 2, 1);
            page1.DrawRectangle(0, 1, 1, 2);
            page1.DrawRectangle(1, 1, 2, 2);

            var active_window = app.ActiveWindow;
            active_window.SelectAll();
            active_window.DeselectAll();

            var selected = active_window.Selection.ToEnumerable().ToDictionary(s => s);
            MUT.Assert.AreEqual(0, selected.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}

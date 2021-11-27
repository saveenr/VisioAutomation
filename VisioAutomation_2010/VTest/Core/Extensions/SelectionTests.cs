using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Core.Extensions
{
    [MUT.TestClass]
    public class SelectionTests : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void Selection_GetShapeIDs()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(0, 1, 1, 2);
            var s3 = page1.DrawRectangle(1, 0, 2, 1);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var application = page1.Application;
            var active_window = application.ActiveWindow;
            active_window.SelectAll();

            var selection = active_window.Selection;
            var selected_ids = selection.GetIDs();

            MUT.Assert.AreEqual(4, selected_ids.Length);
            MUT.Assert.IsNotNull(selected_ids.Contains(s1.ID));
            MUT.Assert.IsNotNull(selected_ids.Contains(s2.ID));
            MUT.Assert.IsNotNull(selected_ids.Contains(s3.ID));
            MUT.Assert.IsNotNull(selected_ids.Contains(s4.ID));
            page1.Delete(1);
        }

        [MUT.TestMethod]
        public void Selection_ToEnumerable()
        {
            // Selection Object: http://msdn.microsoft.com/en-us/library/ms408990(v=office.12).aspx
            // this is a 1-based collection

            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(0, 1, 1, 2);
            var s3 = page1.DrawRectangle(1, 0, 2, 1);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var application = page1.Application;
            var active_window = application.ActiveWindow;
            active_window.SelectAll();

            var selection = active_window.Selection;
            var expected = selection.Cast<IVisio.Shape>().ToList();
            var actual = selection.ToList();

            MUT.Assert.AreEqual(expected.Count, actual.Count);
            MUT.Assert.AreEqual(selection[1].ID16, actual[0].ID16);
            MUT.Assert.AreEqual(selection[2].ID16, actual[1].ID16);
            MUT.Assert.AreEqual(selection[3].ID16, actual[2].ID16);
            MUT.Assert.AreEqual(selection[4].ID16, actual[3].ID16);
            MUT.Assert.AreEqual(selection[1].Index, expected[0].Index);
            MUT.Assert.AreEqual(selection[2].Index, expected[1].Index);
            MUT.Assert.AreEqual(selection[3].Index, expected[2].Index);
            MUT.Assert.AreEqual(selection[4].Index, expected[3].Index);

            MUT.Assert.AreEqual(expected[0].ID16, actual[0].ID16);
            MUT.Assert.AreEqual(expected[1].ID16, actual[1].ID16);
            MUT.Assert.AreEqual(expected[2].ID16, actual[2].ID16);
            MUT.Assert.AreEqual(expected[3].ID16, actual[3].ID16);
            MUT.Assert.AreEqual(expected[0].Index, actual[0].Index);
            MUT.Assert.AreEqual(expected[1].Index, actual[1].Index);
            MUT.Assert.AreEqual(expected[2].Index, actual[2].Index);
            MUT.Assert.AreEqual(expected[3].Index, actual[3].Index);

            page1.Delete(1);
        }
    }
}
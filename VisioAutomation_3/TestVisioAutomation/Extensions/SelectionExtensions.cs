using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class SelectionExtensions : VisioAutomationTest
    {
        [TestMethod]
        public void GetShapeIDs()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(0, 1, 1, 2);
            var s3 = page1.DrawRectangle(1, 0, 2, 1);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            var application = page1.Application;
            var active_window = application.ActiveWindow;
            active_window.SelectAll();

            var selection = active_window.Selection;
            var selected_ids = selection.GetIDs();

            Assert.AreEqual(4, selected_ids.Length);
            Assert.IsNotNull(selected_ids.Contains(s1.ID));
            Assert.IsNotNull(selected_ids.Contains(s2.ID));
            Assert.IsNotNull(selected_ids.Contains(s3.ID));
            Assert.IsNotNull(selected_ids.Contains(s4.ID));
            page1.Delete(1);
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TabStop_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void SetTabStops()
        {
            var no_tab_stops = new VA.Text.TabStop[] { };
            var tabstops = new[]
                               {
                                   new VA.Text.TabStop(0.5, VA.Text.TabStopAlignment.Left),
                                   new VA.Text.TabStop(1.5, VA.Text.TabStopAlignment.Right),
                                   new VA.Text.TabStop(2.5, VA.Text.TabStopAlignment.Center),
                                   new VA.Text.TabStop(4.0, VA.Text.TabStopAlignment.Decimal)
                               };

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            // shapes shoould not have any tab stops by default
            Assert.AreEqual(0, VA.Text.TextHelper.GetTabStopCount(s1));

            // clearing tab stops shoudl work even if there are no tab stops
            VA.Text.TextHelper.SetTabStops(s1, no_tab_stops);
            Assert.AreEqual(0, VA.Text.TextHelper.GetTabStopCount(s1));

            // set the 3 tab stops
            VA.Text.TextHelper.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            Assert.AreEqual(tabstops.Length, VA.Text.TextHelper.GetTabStopCount(s1));

            // clear the tab stops
            VA.Text.TextHelper.SetTabStops(s1, no_tab_stops);
            Assert.AreEqual(0, VA.Text.TextHelper.GetTabStopCount(s1));

            page1.Delete(0);
        }
    }
}
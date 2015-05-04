using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Text;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextFormat_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Text_TabStops_Set()
        {
            var no_tab_stops = new TabStop[] { };
            var tabstops = new[]
                               {
                                   new TabStop(0.5, TabStopAlignment.Left),
                                   new TabStop(1.5, TabStopAlignment.Right),
                                   new TabStop(2.5, TabStopAlignment.Center),
                                   new TabStop(4.0, TabStopAlignment.Decimal)
                               };

            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            // shapes shoould not have any tab stops by default
            var m0 = TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m0.TabStops.Count);

            // clearing tab stops shoudl work even if there are no tab stops
            TextFormat.SetTabStops(s1, no_tab_stops);
            var m1 = TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m1.TabStops.Count);

            // set the 3 tab stops
            TextFormat.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            var m2 = TextFormat.GetFormat(s1);
            Assert.AreEqual(tabstops.Length, m2.TabStops.Count);
            Assert.AreEqual(0.5, tabstops[0].Position);
            Assert.AreEqual(1.5, tabstops[1].Position);
            Assert.AreEqual(2.5, tabstops[2].Position);
            Assert.AreEqual(4.0, tabstops[3].Position);
            Assert.AreEqual(TabStopAlignment.Left, tabstops[0].Alignment);
            Assert.AreEqual(TabStopAlignment.Right, tabstops[1].Alignment);
            Assert.AreEqual(TabStopAlignment.Center, tabstops[2].Alignment);
            Assert.AreEqual(TabStopAlignment.Decimal, tabstops[3].Alignment);

            // clear the tab stops
            TextFormat.SetTabStops(s1, no_tab_stops);
            var m3 = TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m3.TabStops.Count);

            page1.Delete(0);
        }
    }
}
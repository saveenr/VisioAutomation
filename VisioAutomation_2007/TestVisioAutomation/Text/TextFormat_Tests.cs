using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextFormat_Tests : VisioAutomationTest
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
            var m0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m0.TabStops.Count);

            // clearing tab stops shoudl work even if there are no tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m1.TabStops.Count);

            // set the 3 tab stops
            VA.Text.TextFormat.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            var m2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(tabstops.Length, m2.TabStops.Count);
            Assert.AreEqual(0.5, tabstops[0].Position);
            Assert.AreEqual(1.5, tabstops[1].Position);
            Assert.AreEqual(2.5, tabstops[2].Position);
            Assert.AreEqual(4.0, tabstops[3].Position);
            Assert.AreEqual(VA.Text.TabStopAlignment.Left, tabstops[0].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Right, tabstops[1].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Center, tabstops[2].Alignment);
            Assert.AreEqual(VA.Text.TabStopAlignment.Decimal, tabstops[3].Alignment);

            // clear the tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m3 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m3.TabStops.Count);

            page1.Delete(0);
        }
    }
}
using VATEXT = VisioAutomation.Text;

namespace VisioAutomation_Tests.Core.Text
{
    [TestClass]
    public class TextFormatTests : VisioAutomationTest
    {
        [TestMethod]
        public void Text_TabStops_Set()
        {
            var no_tab_stops = new VisioAutomation.Text.TabStop[] { };
            var tabstops = new[]
                               {
                                   new VisioAutomation.Text.TabStop(0.5, VATEXT.TabStopAlignment.Left),
                                   new VisioAutomation.Text.TabStop(1.5, VATEXT.TabStopAlignment.Right),
                                   new VisioAutomation.Text.TabStop(2.5, VATEXT.TabStopAlignment.Center),
                                   new VisioAutomation.Text.TabStop(4.0, VATEXT.TabStopAlignment.Decimal)
                               };

            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            // shapes shoould not have any tab stops by default
            var m0 = VATEXT.TextFormat.GetFormat(s1, VASS.CellValueType.Formula);
            Assert.AreEqual(0, m0.TabStops.Count);

            // clearing tab stops shoudl work even if there are no tab stops
            VATEXT.TextHelper.SetTabStops(s1, no_tab_stops);
            var m1 = VATEXT.TextFormat.GetFormat(s1, VASS.CellValueType.Formula);
            Assert.AreEqual(0, m1.TabStops.Count);

            // set the 3 tab stops
            VATEXT.TextHelper.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            var m2 = VATEXT.TextFormat.GetFormat(s1, VASS.CellValueType.Formula);
            Assert.AreEqual(tabstops.Length, m2.TabStops.Count);
            Assert.AreEqual(0.5, tabstops[0].Position);
            Assert.AreEqual(1.5, tabstops[1].Position);
            Assert.AreEqual(2.5, tabstops[2].Position);
            Assert.AreEqual(4.0, tabstops[3].Position);
            Assert.AreEqual(VATEXT.TabStopAlignment.Left, tabstops[0].Alignment);
            Assert.AreEqual(VATEXT.TabStopAlignment.Right, tabstops[1].Alignment);
            Assert.AreEqual(VATEXT.TabStopAlignment.Center, tabstops[2].Alignment);
            Assert.AreEqual(VATEXT.TabStopAlignment.Decimal, tabstops[3].Alignment);

            // clear the tab stops
            VATEXT.TextHelper.SetTabStops(s1, no_tab_stops);
            var m3 = VATEXT.TextFormat.GetFormat(s1, VASS.CellValueType.Formula);
            Assert.AreEqual(0, m3.TabStops.Count);

            page1.Delete(0);
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VAHLINK = VisioAutomation.Shapes.Hyperlinks;

namespace TestVisioAutomation.Hyperlinks
{
    [TestClass]
    public class Hyperlink_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Hyperlinks_AddRemove()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            // Ensure we start with 0 hyperlinks
            Assert.AreEqual(0, VAHLINK.HyperlinkHelper.GetCount(s1));

            // Add the first hyperlink
            int h1 = VAHLINK.HyperlinkHelper.Add(s1,"http://microsoft.com");
            Assert.AreEqual(1, VAHLINK.HyperlinkHelper.GetCount(s1));

            // Add the second control
            int h2 = VAHLINK.HyperlinkHelper.Add(s1,"http://google.com");
            Assert.AreEqual(2, VAHLINK.HyperlinkHelper.GetCount(s1));
            
            // retrieve the control information
            var hlinks= VAHLINK.HyperlinkCells.GetCells(s1);

            // verify that the hyperlinks were set propery
            Assert.AreEqual(2, hlinks.Count);
            Assert.AreEqual("\"http://microsoft.com\"", hlinks[0].Address.Formula);
            Assert.AreEqual("\"http://google.com\"", hlinks[1].Address.Formula);

            // Delete both hyperlinks
            VAHLINK.HyperlinkHelper.Delete(s1, 0);
            Assert.AreEqual(1, VAHLINK.HyperlinkHelper.GetCount(s1));
            VAHLINK.HyperlinkHelper.Delete(s1, 0);
            Assert.AreEqual(0, VAHLINK.HyperlinkHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void PageOrientation()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Portrait,
                            VA.Pages.PageHelper.GetOrientation(page1));
            Assert.AreEqual(new VA.Drawing.Size(4, 3), VA.Pages.PageHelper.GetSize(page1));

            VA.Pages.PageHelper.SetOrientation(page1, VA.Pages.PrintPageOrientation.Landscape);
            Assert.AreEqual(VA.Pages.PrintPageOrientation.Landscape,
                            VA.Pages.PageHelper.GetOrientation(page1));
            Assert.AreEqual(new VA.Drawing.Size(3, 4), VA.Pages.PageHelper.GetSize(page1));
            page1.Delete(0);
        }

        [TestMethod]
        public void DuplicatePage()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            var s1 = page1.DrawRectangle(1, 1, 3, 3);

            var page2 = VA.Pages.PageHelper.Duplicate(page1, null);

            Assert.AreEqual(new VA.Drawing.Size(4, 3), VA.Pages.PageHelper.GetSize(page2) );
            Assert.AreEqual(1, page2.Shapes.Count);

            var s2 = page2.Shapes[1];
            Assert.AreEqual(new VA.Drawing.Size(2, 2), VisioAutomationTest.GetSize(s2));
            page2.Delete(0);
            page1.Delete(0);
        }
    }
}
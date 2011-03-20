using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class PageHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void ManageBackgroundPagesx()
        {
            // Create the fg page

            var page1 = GetNewPage("1");
            var s1 = page1.DrawRectangle(0, 0, 4, 2);
            s1.Text = "Page1";
            var src_fg = VisioAutomation.ShapeSheet.SRCConstants.FillForegnd;
            var cell_fillfg1 = s1.CellsSRC[src_fg.Section, src_fg.Row, src_fg.Cell ];
            cell_fillfg1.FormulaU = "rgb(250,250,250)";

            // Create the bg page

            var page2 = GetNewPage("2");
            page2.Background = 1;
            var s2 = page2.DrawRectangle(0, 0, 4, 4);
            var cell_fillfg2 = s2.CellsSRC[src_fg.Section, src_fg.Row, src_fg.Cell];
            cell_fillfg2.FormulaU = "rgb(230,180,20)";

            // set the fg to have the other back as a background

            VA.PageHelper.SetBackgroundPage(page1, page2);
            Assert.AreEqual(page2, page1.BackPage);

            // create a new page now

            var page3 = GetNewPage("3");

            // verify that it didn't somehow become a background page because of what we did earlier

            Assert.AreEqual(0, page3.Background);

            // Unassign the bg page from the fg page

            VA.PageHelper.SetBackgroundPage(page1, null);

            // clean-up - delete all the pages

            page3.Delete(0);
            page2.Delete(0);
            page1.Delete(0);
        }

        [TestMethod]
        public void PageOrientation()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            Assert.AreEqual(VA.Layout.PrintPageOrientation.Portrait,
                            VA.PageHelper.GetPageOrientation(page1));
            Assert.AreEqual(new VA.Drawing.Size(4, 3), page1.GetSize());

            VA.PageHelper.SetPageOrientation(page1, VA.Layout.PrintPageOrientation.Landscape);
            Assert.AreEqual(VA.Layout.PrintPageOrientation.Landscape,
                            VA.PageHelper.GetPageOrientation(page1));
            Assert.AreEqual(new VA.Drawing.Size(3, 4), page1.GetSize());
            page1.Delete(0);
        }

        [TestMethod]
        public void DuplicatePage()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(4, 3));
            var s1 = page1.DrawRectangle(1, 1, 3, 3);

            var page2 = VA.PageHelper.Duplicate(page1);

            Assert.AreEqual(new VA.Drawing.Size(4, 3), page2.GetSize() );
            Assert.AreEqual(1, page2.Shapes.Count);

            var s2 = page2.Shapes[1];
            Assert.AreEqual(new VA.Drawing.Size(2, 2), VisioAutomationTest.GetSize(s2));
            page2.Delete(0);
            page1.Delete(0);
        }
    }
}
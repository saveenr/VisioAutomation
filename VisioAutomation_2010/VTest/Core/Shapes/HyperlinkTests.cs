using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class HyperlinkTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void GetCount_OnFreshShape_ReturnsZero()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            MUT.Assert.AreEqual(0, VA.Shapes.HyperlinkHelper.GetCount(s1));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Add_TwoHyperlinksToShape_GetCellsReturnsBothAddressesInOrder()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            var h1 = new VA.Shapes.HyperlinkCells();
            h1.Address = "http://microsoft.com";
            VA.Shapes.HyperlinkHelper.Add(s1, h1);
            MUT.Assert.AreEqual(1, VA.Shapes.HyperlinkHelper.GetCount(s1));

            var h2 = new VA.Shapes.HyperlinkCells();
            h2.Address = "http://google.com";
            VA.Shapes.HyperlinkHelper.Add(s1, h2);
            MUT.Assert.AreEqual(2, VA.Shapes.HyperlinkHelper.GetCount(s1));

            var hlinks = VA.Shapes.HyperlinkCells.GetCells(s1, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(2, hlinks.Count);
            MUT.Assert.AreEqual("\"http://microsoft.com\"", hlinks[0].Address.Value);
            MUT.Assert.AreEqual("\"http://google.com\"", hlinks[1].Address.Value);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Delete_RemovesHyperlinksOneAtATime_CountDropsToZero()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            var h1 = new VA.Shapes.HyperlinkCells();
            h1.Address = "http://microsoft.com";
            VA.Shapes.HyperlinkHelper.Add(s1, h1);
            var h2 = new VA.Shapes.HyperlinkCells();
            h2.Address = "http://google.com";
            VA.Shapes.HyperlinkHelper.Add(s1, h2);
            MUT.Assert.AreEqual(2, VA.Shapes.HyperlinkHelper.GetCount(s1));

            VA.Shapes.HyperlinkHelper.Delete(s1, 0);
            MUT.Assert.AreEqual(1, VA.Shapes.HyperlinkHelper.GetCount(s1));
            VA.Shapes.HyperlinkHelper.Delete(s1, 0);
            MUT.Assert.AreEqual(0, VA.Shapes.HyperlinkHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}

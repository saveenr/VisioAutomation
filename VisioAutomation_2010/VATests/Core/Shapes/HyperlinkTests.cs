namespace VisioAutomation_Tests.Core.Shapes;

[TestClass]
public class HyperlinkTests : VisioAutomationTest
{
    [TestMethod]
    public void Hyperlinks_AddRemove()
    {
        var page1 = this.GetNewPage();

        var s1 = page1.DrawRectangle(0, 0, 4, 1);

        // Ensure we start with 0 hyperlinks
        Assert.AreEqual(0, VA.Shapes.HyperlinkHelper.GetCount(s1));

        // Add the first hyperlink

        var h1 = new VA.Shapes.HyperlinkCells();
        h1.Address = "http://microsoft.com";
        int h1_row = VA.Shapes.HyperlinkHelper.Add(s1,h1);
        Assert.AreEqual(1, VA.Shapes.HyperlinkHelper.GetCount(s1));

        // Add the second control
        var h2 = new VA.Shapes.HyperlinkCells();
        h2.Address = "http://google.com";
        int h2_row = VA.Shapes.HyperlinkHelper.Add(s1,h2);
        Assert.AreEqual(2, VA.Shapes.HyperlinkHelper.GetCount(s1));
            
        // retrieve the control information
        var hlinks= VA.Shapes.HyperlinkCells.GetCells(s1, VASS.CellValueType.Formula);

        // verify that the hyperlinks were set propery
        Assert.AreEqual(2, hlinks.Count);
        Assert.AreEqual("\"http://microsoft.com\"", hlinks[0].Address.Value);
        Assert.AreEqual("\"http://google.com\"", hlinks[1].Address.Value);

        // Delete both hyperlinks
        VA.Shapes.HyperlinkHelper.Delete(s1, 0);
        Assert.AreEqual(1, VA.Shapes.HyperlinkHelper.GetCount(s1));
        VA.Shapes.HyperlinkHelper.Delete(s1, 0);
        Assert.AreEqual(0, VA.Shapes.HyperlinkHelper.GetCount(s1));

        page1.Delete(0);
    }
}
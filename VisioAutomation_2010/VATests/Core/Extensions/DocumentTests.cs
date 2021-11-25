using VisioAutomation.Extensions;

namespace VisioAutomation_Tests.Core.Extensions;

[TestClass]
public class DocumentTests : VisioAutomationTest
{
    [TestMethod]
    public void Document_ForceClose()
    {
        var app = this.GetVisioApplication();
        var documents = app.Documents;
        int old_count = documents.Count;
        var doc1 = documents.Add(string.Empty);
        Assert.AreEqual(old_count + 1, documents.Count);
        var page1 = doc1.Pages[1];
        var s1 = page1.DrawRectangle(1, 1, 2, 2);
        doc1.Close(true);
        Assert.AreEqual(old_count, documents.Count);
    }
}
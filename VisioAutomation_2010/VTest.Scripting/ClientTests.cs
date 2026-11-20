using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class ClientTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Document_CloseAllDocumentsWithoutSaving_NewDocStartsWithNoSelection()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllDocumentsWithoutSaving();

            MUT.Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.NewDocument();
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 2, 3, 3);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
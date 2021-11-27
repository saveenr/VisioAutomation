using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class ScriptingDocumentTests : VTest.VisioAutomationTest
    {
        [MUT.TestMethod]
        public void Document_Activation()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            var doc1 = client.Document.NewDocument();
            var doc2 = client.Document.NewDocument();
            var doc3 = client.Document.NewDocument();

            client.Document.ActivateDocument(doc1);
            MUT.Assert.AreEqual(doc1, app.ActiveDocument);
            client.Document.ActivateDocument(doc2);
            MUT.Assert.AreEqual(doc2, app.ActiveDocument);
            client.Document.ActivateDocument(doc3);
            MUT.Assert.AreEqual(doc3, app.ActiveDocument);
            client.Document.ActivateDocument(doc1);
            MUT.Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }
    }
}
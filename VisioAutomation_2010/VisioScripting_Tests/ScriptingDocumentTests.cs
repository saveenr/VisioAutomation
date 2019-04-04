using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingDocumentTests : VisioAutomationTest
    {
        [TestMethod]
        public void Document_Activation()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetAttachedApplication();
            var doc1 = client.Document.NewDocument();
            var doc2 = client.Document.NewDocument();
            var doc3 = client.Document.NewDocument();

            client.Document.ActivateDocument(doc1);
            Assert.AreEqual(doc1, app.ActiveDocument);
            client.Document.ActivateDocument(doc2);
            Assert.AreEqual(doc2, app.ActiveDocument);
            client.Document.ActivateDocument(doc3);
            Assert.AreEqual(doc3, app.ActiveDocument);
            client.Document.ActivateDocument(doc1);
            Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }
    }
}
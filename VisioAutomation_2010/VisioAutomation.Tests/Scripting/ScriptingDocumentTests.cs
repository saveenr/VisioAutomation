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
            var app = client.Application.Get();
            var doc1 = client.Document.New();
            var doc2 = client.Document.New();
            var doc3 = client.Document.New();

            client.Document.Activate(doc1);
            Assert.AreEqual(doc1, app.ActiveDocument);
            client.Document.Activate(doc2);
            Assert.AreEqual(doc2, app.ActiveDocument);
            client.Document.Activate(doc3);
            Assert.AreEqual(doc3, app.ActiveDocument);
            client.Document.Activate(doc1);
            Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }
    }
}
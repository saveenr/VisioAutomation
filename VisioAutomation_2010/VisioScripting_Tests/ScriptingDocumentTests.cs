using UT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;

namespace VisioScripting_Tests
{
    [UT.TestClass]
    public class ScriptingDocumentTests : VisioAutomation_Tests.VisioAutomationTest
    {
        [UT.TestMethod]
        public void Document_Activation()
        {
            var client = this.GetScriptingClient();
            var app = client.Application.GetApplication();
            var doc1 = client.Document.NewDocument();
            var doc2 = client.Document.NewDocument();
            var doc3 = client.Document.NewDocument();

            client.Document.ActivateDocument(doc1);
            UT.Assert.AreEqual(doc1, app.ActiveDocument);
            client.Document.ActivateDocument(doc2);
            UT.Assert.AreEqual(doc2, app.ActiveDocument);
            client.Document.ActivateDocument(doc3);
            UT.Assert.AreEqual(doc3, app.ActiveDocument);
            client.Document.ActivateDocument(doc1);
            UT.Assert.AreEqual(doc1, app.ActiveDocument);

            doc1.Close(true);
            doc2.Close(true);
            doc3.Close(true);
        }
    }
}
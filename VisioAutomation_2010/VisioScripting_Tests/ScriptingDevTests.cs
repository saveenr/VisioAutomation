using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioScripting_Tests
{
    [TestClass]
    public class ScriptingDevTests : VisioAutomation_Tests.VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Dev_ScriptingDocumentation()
        {
            var client = this.GetScriptingClient();
            client.Developer.DrawScriptingDocumentation();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
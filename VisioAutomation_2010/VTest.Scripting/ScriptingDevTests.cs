using UT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioScripting_Tests
{
    [UT.TestClass]
    public class ScriptingDevTests : VTest.VisioAutomationTest
    {
        [UT.TestMethod]
        public void Scripting_Dev_ScriptingDocumentation()
        {
            var client = this.GetScriptingClient();
            client.Developer.DrawScriptingDocumentation();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
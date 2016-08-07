using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingDevTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Dev_ScriptingDocumentation()
        {
            var client = this.GetScriptingClient();
            client.Developer.DrawScriptingDocumentation();
            client.Document.Close(true);
        }
    }
}
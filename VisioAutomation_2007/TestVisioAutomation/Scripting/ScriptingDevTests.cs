using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingDevTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Dev_ScriptingDocumentation()
        {
            var ss = GetScriptingClient();
            ss.Developer.DrawScriptingDocumentation();
            ss.Document.Close(true);
        }
    }
}
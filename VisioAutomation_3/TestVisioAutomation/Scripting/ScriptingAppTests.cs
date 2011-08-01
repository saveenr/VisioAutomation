using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingSessionTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_GetWindowText()
        {
            var ss = GetScriptingSession();

            ss.Developer.DrawDocumentation();

        }

    }
}
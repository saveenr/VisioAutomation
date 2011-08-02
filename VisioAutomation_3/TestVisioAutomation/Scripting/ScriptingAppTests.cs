using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingSessionTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_DevDocumentation()
        {
            var ss = GetScriptingSession();
            var doc= ss.Developer.DrawDocumentation();
            doc.Close(true);
        }
    }
}
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

            string t1 = ss.Document.GetHelp();
            string t2 = ss.CustomProp.GetHelp();


        }

    }
}
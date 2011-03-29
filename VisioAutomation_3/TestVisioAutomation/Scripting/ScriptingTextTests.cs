using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingTextTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Text_ToggleCase_1()
        {
            var scriptingsession = GetScriptingSession();
            scriptingsession.Document.NewDocument();
            scriptingsession.Page.NewPage(new VA.Drawing.Size(4, 4), false);

            var shape_rect = scriptingsession.Draw.DrawRectangle(1, 1, 3, 3);

            scriptingsession.Text.SetText("Hello World");
            Assert.AreEqual("Hello World", scriptingsession.Text.GetText()[0]);

            scriptingsession.Text.ToogleCase();
            Assert.AreEqual("HELLO WORLD", scriptingsession.Text.GetText()[0]);

            scriptingsession.Text.ToogleCase();
            Assert.AreEqual("hello world", scriptingsession.Text.GetText()[0]);

            scriptingsession.Text.ToogleCase();
            Assert.AreEqual("Hello World", scriptingsession.Text.GetText()[0]);

            scriptingsession.Document.CloseDocument(true);
        }
    }
}
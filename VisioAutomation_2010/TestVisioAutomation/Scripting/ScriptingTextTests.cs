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
            scriptingsession.Document.New();
            scriptingsession.Page.New(new VA.Drawing.Size(4, 4), false);

            var shape_rect = scriptingsession.Draw.Rectangle(1, 1, 3, 3);

            shape_rect.Text = "Hello World";
            var t0 = scriptingsession.Text.GetText()[0];
            scriptingsession.Text.SetText("Hello World");
            Assert.AreEqual("Hello World", t0);

            scriptingsession.Text.ToogleCase();
            var t1 = scriptingsession.Text.GetText()[0];
            Assert.AreEqual("HELLO WORLD", t1);

            scriptingsession.Text.ToogleCase();
            var t2 = scriptingsession.Text.GetText()[0];
            Assert.AreEqual("hello world", t2);

            scriptingsession.Text.ToogleCase();
            var t3 = scriptingsession.Text.GetText()[0];
            Assert.AreEqual("Hello World", t3);

            scriptingsession.Document.Close(true);
        }
    }
}
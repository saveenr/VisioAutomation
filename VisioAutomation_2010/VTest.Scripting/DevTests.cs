using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class DevTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Dev_ScriptingDocumentation()
        {
            var client = this.GetScriptingClient();
            client.Developer.DrawScriptingDocumentation();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
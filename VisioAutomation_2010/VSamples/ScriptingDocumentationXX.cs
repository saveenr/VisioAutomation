namespace VSamples
{
    public static class ScriptingDocumentationXX
    {
        public static void ScriptingDocumentation()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawScriptingDocumentation();
        }
    }
}
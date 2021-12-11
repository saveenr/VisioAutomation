namespace VSamples
{
    public static class DeveloperSamples1
    {
        public static void ScriptingDocumentation()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawScriptingDocumentation();
        }
    }
}
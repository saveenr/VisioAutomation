namespace VisioAutomationSamples
{
    public static class DeveloperSamples
    {
        public static void ScriptingDocumentation()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioAutomation.Scripting.Client(app);
            var doc = client.Developer.DrawScriptingDocumentation();
        }

        public static void InteropEnumDocumentation()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioAutomation.Scripting.Client(app);
            var doc = client.Developer.DrawInteropEnumDocumentation();
        }

        public static void VisioAutomationNamespaces()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioAutomation.Scripting.Client(app);
            var doc = client.Developer.DrawNamespaces();
        }

        public static void VisioAutomationNamespacesAndClasses()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioAutomation.Scripting.Client(app);
            var doc = client.Developer.DrawNamespacesAndClasses();
        }
    }
}
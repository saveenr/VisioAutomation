namespace VSamples
{

    public static class DeveloperSamples3
    {

        public static void VisioAutomationNamespaces()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawNamespaces();
        }
    }
}
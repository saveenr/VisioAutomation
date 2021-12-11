namespace VSamples
{
    public static class DeveloperSamples2
    {


        public static void InteropEnumDocumentation()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawInteropEnumDocumentation();
        }
    }
}
namespace VSamples.Samples.Developer
{
    public class DocumentScriptingAPI : SampleMethodBase
    {
        public override  void Run()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawScriptingDocumentation();
        }
    }
}
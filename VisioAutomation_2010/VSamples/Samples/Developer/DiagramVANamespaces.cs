namespace VSamples.Samples.Developer
{

    public  class DiagramVANamespaces : SampleMethodBase
    {

        public override void Run()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawNamespaces();
        }
    }
}
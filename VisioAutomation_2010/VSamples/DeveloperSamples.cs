namespace VSamples
{
    public  class VisioAutomationNamespacesAndClassesX : SampleMethodBase
    {

        public override void RunSample()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawNamespacesAndClasses();
        }
    }
}
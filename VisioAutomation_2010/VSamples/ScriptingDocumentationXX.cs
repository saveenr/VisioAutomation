﻿namespace VSamples
{
    public class ScriptingDocumentationXX : SampleMethodBase
    {
        public override  void RunSample()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawScriptingDocumentation();
        }
    }
}
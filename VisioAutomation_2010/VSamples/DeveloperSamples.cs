﻿namespace VSamples
{
    public static class DeveloperSamples4
    {

        public static void VisioAutomationNamespacesAndClasses()
        {
            var app = SampleEnvironment.Application;
            var client = new VisioScripting.Client(app);
            var doc = client.Developer.DrawNamespacesAndClasses();
        }
    }
}
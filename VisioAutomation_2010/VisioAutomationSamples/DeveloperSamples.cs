using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class DeveloperSamples
    {
        public static void ScriptingDocumentation()
        {
            var app = SampleEnvironment.Application;
            var ss = new VisioAutomation.Scripting.Session(app);
            var doc = ss.Developer.DrawScriptingDocumentation();
        }

        public static void InteropEnumDocumentation()
        {
            var app = SampleEnvironment.Application;
            var ss = new VisioAutomation.Scripting.Session(app);
            var doc = ss.Developer.DrawInteropEnumDocumentation();
        }

        public static void VisioAutomationNamespaces()
        {
            var app = SampleEnvironment.Application;
            var ss = new VisioAutomation.Scripting.Session(app);
            var doc = ss.Developer.DrawVANamespaces();
        }

        public static void VisioAutomationNamespacesAndClasses()
        {
            var app = SampleEnvironment.Application;
            var ss = new VisioAutomation.Scripting.Session(app);
            var doc = ss.Developer.DrawVANamespacesAndClasses();
        }

    }
}
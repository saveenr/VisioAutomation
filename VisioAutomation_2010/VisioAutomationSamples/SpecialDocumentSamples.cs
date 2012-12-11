using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using OCMODEL = VisioAutomation.Layout.Models.OrgChart;

namespace VisioAutomationSamples
{
    public static class SpecialDocumentSamples
    {
        public static void OrgChart()
        {
            // This creates a new document
            var orgchart = new OCMODEL.Drawing();

            var bob = new OCMODEL.Node("Bob");
            var ted = new OCMODEL.Node("Ted");
            var alice = new OCMODEL.Node("Alice");

            bob.Children.Add(ted);
            bob.Children.Add(alice);

            orgchart.Root = bob;

            orgchart.Render(SampleEnvironment.Application);

            var bordersize = new VA.Drawing.Size(1,1);
            SampleEnvironment.Application.ActivePage.ResizeToFitContents(bordersize);
        }
    }
}
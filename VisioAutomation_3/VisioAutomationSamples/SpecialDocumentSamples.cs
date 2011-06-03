using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class SpecialDocumentSamples
    {
        public static void OrgChart()
        {
            // This creates a new document
            var orgchart = new VA.Layout.OrgChart.Drawing();

            var bob = new VA.Layout.OrgChart.Node("Bob");
            var ted = new VA.Layout.OrgChart.Node("Ted");
            var alice = new VA.Layout.OrgChart.Node("Alice");

            bob.Children.Add(ted);
            bob.Children.Add(alice);

            orgchart.Root = bob;

            orgchart.Render(SampleEnvironment.Application);

            SampleEnvironment.Application.ActivePage.ResizeToFitContents(1,1);
        }
    }
}
using OCMODEL = VisioAutomation.Models.Documents.OrgCharts;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples;

public static class SpecialDocumentSamples
{
    public static void OrgChart()
    {
        // This creates a new document
        var orgchart = new OCMODEL.OrgChartDocument();

        var bob = new OCMODEL.Node("Bob");
        var ted = new OCMODEL.Node("Ted");
        var alice = new OCMODEL.Node("Alice");

        bob.Children.Add(ted);
        bob.Children.Add(alice);

        orgchart.OrgCharts.Add(bob);

        orgchart.Render(SampleEnvironment.Application);

        var bordersize = new VA.Geometry.Size(1,1);
        SampleEnvironment.Application.ActivePage.ResizeToFitContents(bordersize);
    }
}
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using OCMODEL = VisioAutomation.Models.Documents.OrgCharts;

namespace VSamples.Samples.Misc
{
    public class OrgChart1 : SampleMethodBase
    {
        public override void RunSample()
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

            var bordersize = new VA.Core.Size(1, 1);
            SampleEnvironment.Application.ActivePage.ResizeToFitContents(bordersize);
        }
    }
}
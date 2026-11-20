using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VAORGCHART = VisioAutomation.Models.Documents.OrgCharts;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page models/org-charts.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ModelsOrgChartsSamples
    {
        public static void Sample_01(IVisio.Application visioApp)
        {
            // <sample id="models/org-charts#01" doc="models/org-charts.md" line="24">
            var orgchart = new VAORGCHART.OrgChartDocument();

            var ceo = new VAORGCHART.Node("Alex (CEO)");
            ceo.Size = new VA.Core.Size(4, 2);
            orgchart.OrgCharts.Add(ceo);

            orgchart.Render(visioApp);
            // </sample>
        }

        public static void Sample_02(IVisio.Application visioApp, VAORGCHART.OrgChartDocument orgchart)
        {
            // <sample id="models/org-charts#02" doc="models/org-charts.md" line="45">
            var ceo = new VAORGCHART.Node("Alex (CEO)");
            ceo.Size = new VA.Core.Size(4, 2);

            var cto = new VAORGCHART.Node("Sam (CTO)");
            var cfo = new VAORGCHART.Node("Robin (CFO)");

            var lead1 = new VAORGCHART.Node("Casey (Lead)");
            var lead2 = new VAORGCHART.Node("Pat (Lead)");

            ceo.Children.Add(cto);
            ceo.Children.Add(cfo);
            cto.Children.Add(lead1);
            cto.Children.Add(lead2);

            orgchart.OrgCharts.Add(ceo);
            orgchart.Render(visioApp);
            // </sample>
        }

        public static void Sample_03()
        {
            // <sample id="models/org-charts#03" doc="models/org-charts.md" line="72">
            var lead1 = new VAORGCHART.Node("Casey (Lead)");
            lead1.Url = "https://example.internal/people/casey";
            // </sample>
        }

        public static void Sample_04(VAORGCHART.OrgChartDocument orgchart)
        {
            // <sample id="models/org-charts#04" doc="models/org-charts.md" line="81">
            orgchart.OrgChartLayoutOptions.Direction = VAORGCHART.OrgChartLayoutDirection.Right;
            // </sample>
        }

        public static void Sample_05(IVisio.Page page, IVisio.Application visioApp)
        {
            // <sample id="models/org-charts#05" doc="models/org-charts.md" line="118">
            var orgchart = new VAORGCHART.OrgChartDocument();

            var team_a = new VAORGCHART.Node("A");
            team_a.Children.Add(new VAORGCHART.Node("B"));

            var team_x = new VAORGCHART.Node("X");
            team_x.Children.Add(new VAORGCHART.Node("Y"));

            orgchart.OrgCharts.Add(team_a);   // page 1
            orgchart.OrgCharts.Add(team_x);   // page 2

            orgchart.Render(visioApp);
            // </sample>
        }
    }
}

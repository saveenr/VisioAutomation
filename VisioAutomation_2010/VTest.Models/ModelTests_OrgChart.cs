using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using SXL = System.Xml.Linq;

namespace VTest.Models
{
    public partial class ModelTests : Framework.VTest
    {
        [MUT.TestMethod]
        [MUT.DeploymentItem(@"datafiles\orgchart_1.xml", "datafiles")]
        public void Scripting_Draw_OrgChart()
        {
            // Load the chart
            string xml = this.get_datafile_content(@"datafiles\orgchart_1.xml");
            
            // Draw the Chart
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            this.draw_org_chart(client, xml);

            // Cleanup
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        private void draw_org_chart(VisioScripting.Client client, string text)
        {
            var xmldoc = SXL.XDocument.Parse(text);
            var orgchart = VisioScripting.Builders.OrgChartDocumentLoader.LoadFromXml(client, xmldoc);

            client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, orgchart);
        }

    }
}
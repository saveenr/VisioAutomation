using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioModelOrgChart)]
    public class NewVisioModelOrgChart : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
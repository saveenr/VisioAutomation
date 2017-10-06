using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelOrgChart)]
    public class NewVisioModelOrgChart : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new VA.Models.Documents.OrgCharts.OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
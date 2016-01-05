using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioOrgChartModel)]
    public class New_VisioOrgChartModel : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new VA.Models.OrgChart.OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
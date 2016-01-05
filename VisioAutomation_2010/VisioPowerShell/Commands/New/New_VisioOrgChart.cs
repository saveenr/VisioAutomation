using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioOrgChart)]
    public class New_VisioOrgChart : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new VA.Models.OrgChart.OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
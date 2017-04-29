using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelOrgChart)]
    public class New_VisioModelOrgChart : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new VA.Models.Documents.OrgCharts.OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
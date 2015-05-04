using VisioAutomation.Models.OrgChart;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioOrgChart")]
    public class New_VisioOrgChart : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var orgchart = new OrgChartDocument();
            this.WriteObject(orgchart);
        }
    }
}
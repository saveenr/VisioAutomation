using VisioPS.Extensions;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "OrgChart")]
    public class New_OrgChart : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var oc = new VA.Layout.OrgChart.Drawing();
            this.WriteObject(oc);
        }
    }
}
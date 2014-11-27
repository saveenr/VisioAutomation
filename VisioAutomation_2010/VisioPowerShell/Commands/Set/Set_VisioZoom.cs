using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioZoom")]
    public class Set_VisioZoom : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public VA.Scripting.Zoom Level = VA.Scripting.Zoom.In;

        [SMA.Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
        public double Percent = 0;

        protected override void ProcessRecord()
        {
            if (this.Percent > 0)
            {
                this.client.View.ZoomToPercentage(this.Percent);
            }
            else
            {
                this.client.View.Zoom(this.Level);       
            }
        }
    }
}
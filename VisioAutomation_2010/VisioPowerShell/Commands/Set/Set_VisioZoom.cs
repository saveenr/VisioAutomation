using System.Management.Automation;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioZoom")]
    public class Set_VisioZoom : VisioCmdlet
    {
        [Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public VisioAutomation.Scripting.Zoom Level = VisioAutomation.Scripting.Zoom.In;

        [Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
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
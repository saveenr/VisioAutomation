using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioZoom)]
    public class SetVisioZoom : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public VisioScripting.Models.Zoom Level = VisioScripting.Models.Zoom.In;

        [SMA.Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
        public double Percent = 0;

        protected override void ProcessRecord()
        {
            if (this.Percent > 0)
            {
                this.Client.View.ZoomToPercentage(this.Percent);
            }
            else
            {
                this.Client.View.Zoom(this.Level);       
            }
        }
    }
}
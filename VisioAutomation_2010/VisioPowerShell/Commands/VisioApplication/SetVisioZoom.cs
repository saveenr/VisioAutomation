using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioZoom)]
    public class SetVisioZoom : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public VisioScripting.Models.Zoom? Level = null;

        [SMA.Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
        public double Percent = 0;

        [SMA.Parameter(ParameterSetName = "relativepercent", Position = 0, Mandatory = true)]
        public double PercentMultiplier = 0;

        protected override void ProcessRecord()
        {
            if (this.Percent > 0)
            {
                this.Client.View.SetActiveWindowToZoom(this.Percent);
            }
            else if (this.PercentMultiplier > 0)
            {
                this.Client.View.ZoomActiveWindowRelative(this.PercentMultiplier);
            }
            else if (this.Level != null)
            {
                this.Client.View.ZoomActiveWindowToObject(this.Level.Value);       
            }
            else
            {
                throw new System.ArgumentException("Must provide a parameter");
            }
        }
    }
}
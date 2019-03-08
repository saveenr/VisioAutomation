using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioZoom)]
    public class SetVisioZoom : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "zoomto", Position = 0, Mandatory = true)] 
        public VisioScripting.Models.ZoomToObject? To = null;

        [SMA.Parameter(ParameterSetName = "value", Position = 0, Mandatory = true)] 
        public double Value = 0;

        [SMA.Parameter(ParameterSetName = "valuerelative", Position = 0, Mandatory = true)]
        public double ValueRelative = 0;

        protected override void ProcessRecord()
        {
            if (this.Value > 0)
            {
                this.Client.View.SetActiveWindowZoomValue(this.Value);
            }
            else if (this.ValueRelative > 0)
            {
                this.Client.View.SetActiveWindowZoomValueRelative(this.ValueRelative);
            }
            else if (this.To != null)
            {
                this.Client.View.SetActiveWindowZoomToObject(this.To.Value);       
            }
            else
            {
                throw new System.ArgumentException("Must provide a parameter");
            }
        }
    }
}
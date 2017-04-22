using System.Management.Automation;
using VisioScripting.Models;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioZoom)]
    public class Set_VisioZoom : VisioCmdlet
    {
        [Parameter(ParameterSetName = "level", Position = 0, Mandatory = true)] 
        public Zoom Level = Zoom.In;

        [Parameter(ParameterSetName = "percent", Position = 0, Mandatory = true)] 
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
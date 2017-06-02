using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioLayer)]
    public class GetVisioLayer : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            if (this.Name!=null || this.Name=="*")
            {
                var layer = this.Client.Layer.Get(this.Name);
                this.WriteObject(layer);
            }
            else
            {
                var layers = this.Client.Layer.Get();
                this.WriteObject(layers,false);
            }
        }
    }
}
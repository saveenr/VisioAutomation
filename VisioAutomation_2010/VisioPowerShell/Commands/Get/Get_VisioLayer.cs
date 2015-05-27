using System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioLayer")]
    public class Get_VisioLayer : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            if (this.Name!=null || this.Name=="*")
            {
                var layer = this.client.Layer.Get(this.Name);
                this.WriteObject(layer);
            }
            else
            {
                var layers = this.client.Layer.Get();
                this.WriteObject(layers,false);
            }
        }
    }
}
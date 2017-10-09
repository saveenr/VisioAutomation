using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioLayer)]
    public class GetVisioLayer : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            string str_asterisk = "*";
            if (this.Name!=null || this.Name==str_asterisk)
            {
                var layer = this.Client.Layer.Get(this.Name);
                this.WriteObject(layer);
            }
            else
            {
                var layers = this.Client.Layer.Get();
                this.WriteObject(layers, true);
            }
        }
    }
}
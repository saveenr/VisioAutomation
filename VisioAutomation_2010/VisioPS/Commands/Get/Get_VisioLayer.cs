using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioLayer")]
    public class Get_VisioLayer : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Name==null)
            {
                var layer = scriptingsession.Layer.Get(this.Name);
                this.WriteObject(layer);
            }
            else
            {
                var layers = scriptingsession.Layer.Get();
                this.WriteObject(layers,false);
            }
        }
    }
}
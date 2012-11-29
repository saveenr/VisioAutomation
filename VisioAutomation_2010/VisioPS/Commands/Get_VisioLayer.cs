using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "VisioLayer")]
    public class Get_VisioLayer : VisioPSCmdlet
    {

        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Name==null)
            {
                var layer = scriptingsession.Layer.GetLayer(this.Name);
                this.WriteObject(layer);
            }
            else
            {
                var layers = scriptingsession.Layer.GetLayers();
                this.WriteObject(layers);
            }
        }
    }
}
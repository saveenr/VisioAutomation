using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Get", "Edge")]
    public class Get_Edge : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var edges = scriptingsession.Connection.GetEdges();
            this.WriteObject(edges);
        }
    }
}
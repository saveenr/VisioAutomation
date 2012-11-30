using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioEdge")]
    public class Get_VisioEdge : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var edges = scriptingsession.Connection.GetEdges();
            this.WriteObject(edges);
        }
    }
}
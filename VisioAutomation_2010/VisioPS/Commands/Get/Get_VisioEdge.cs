using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioEdge")]
    public class Get_VisioEdge : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public SMA.SwitchParameter GetShapeObjects { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var edges = scriptingsession.Connection.GetEdges();
            if (this.GetShapeObjects)
            {
                this.WriteObject(edges, false);                
            }
            else
            {
                foreach (var edge in edges)
                {
                    var e = new DirectedEdge();
                    e.FromShapeID = edge.From.ID;
                    e.ToShapeID = edge.To.ID;
                    e.ConnectorID = edge.Connector.ID;
                    this.WriteObject(e);      
                }
            }
        }
    }
}
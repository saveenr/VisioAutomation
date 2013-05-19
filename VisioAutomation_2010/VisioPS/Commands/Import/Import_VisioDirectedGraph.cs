using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Import, "VisioDirectedGraph")]
    public class Import_VisioDirectedGraph : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public string Filename { get; set; }

        protected override void ProcessRecord()
        {
            if (!this.CheckFileExists(this.Filename))
            {
                return;
            }

            var scriptingsession = this.ScriptingSession;
            var dg_model = VA.Scripting.DirectedGraph.DirectedGraphBuilder.LoadFromXML(
                scriptingsession, 
                this.Filename);            
            
            this.WriteObject(dg_model);
        }
    }
}
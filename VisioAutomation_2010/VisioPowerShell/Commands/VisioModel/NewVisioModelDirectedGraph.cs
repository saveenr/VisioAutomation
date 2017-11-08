using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioModelDirectedGraph)]
    public class NewVisioModelDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphLayout();
            this.WriteObject(dg_model);
        }
    }
}
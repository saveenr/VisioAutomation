using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioDirectedGraph")]
    public class New_VisioDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new VA.Models.DirectedGraph.Drawing();
            this.WriteObject(dg_model);
        }
    }
}
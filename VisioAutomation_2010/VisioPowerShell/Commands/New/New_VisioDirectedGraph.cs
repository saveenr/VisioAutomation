using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioDirectedGraph")]
    public class New_VisioDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new VA.Models.DirectedGraph.Drawing();
            this.WriteObject(dg_model);
        }
    }
}
using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioModelDirectedGraph)]
    public class New_VisioModelDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new VA.Models.Layouts.DirectedGraph.DirectedGraphLayout();
            this.WriteObject(dg_model);
        }
    }
}
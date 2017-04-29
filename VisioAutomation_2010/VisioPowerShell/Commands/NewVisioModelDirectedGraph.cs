using System.Management.Automation;
using VisioAutomation.Models.Layouts.DirectedGraph;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelDirectedGraph)]
    public class NewVisioModelDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new DirectedGraphLayout();
            this.WriteObject(dg_model);
        }
    }
}
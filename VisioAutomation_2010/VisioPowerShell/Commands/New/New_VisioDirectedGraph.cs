using VisioAutomation.Models.DirectedGraph;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioDirectedGraph")]
    public class New_VisioDirectedGraph : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new Drawing();
            this.WriteObject(dg_model);
        }
    }
}
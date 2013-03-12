using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioDirectedGraph")]
    public class New_VisioDirectedGraph : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var dg_model = new VA.Layout.Models.DirectedGraph.Drawing();           
            this.WriteObject(dg_model);
        }
    }
}
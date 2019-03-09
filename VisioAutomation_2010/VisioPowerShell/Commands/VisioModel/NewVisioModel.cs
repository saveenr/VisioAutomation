using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioModel)]
    public class NewVisioModel: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public VisioPowerShell.Models.DiagramModelType Type = VisioPowerShell.Models.DiagramModelType.DirectedGraph;
        protected override void ProcessRecord()
        {
            if (this.Type == VisioPowerShell.Models.DiagramModelType.DirectedGraph)
            {
                var dg_model = new VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphLayout();
                this.WriteObject(dg_model);
            }
            else if (this.Type == VisioPowerShell.Models.DiagramModelType.OrgChart)
            {
                var orgchart = new VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument();
                this.WriteObject(orgchart);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(this.Type));
            }
        }
    }
}
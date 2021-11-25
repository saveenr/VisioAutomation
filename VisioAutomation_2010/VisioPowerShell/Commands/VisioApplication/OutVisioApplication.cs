using VisioAutomation.Models.Layouts.DirectedGraph;
using MODELS = VisioAutomation.Models;

namespace VisioPowerShell.Commands.VisioApplication;

[SMA.Cmdlet(SMA.VerbsData.Out, Nouns.VisioApplication)]
public class OutVisio : VisioCmdlet
{
    [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "orgchcart")]
    public MODELS.Documents.OrgCharts.OrgChartDocument OrgChart { get; set; }

    [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "grid")]
    public MODELS.Layouts.Grid.GridLayout GridLayout { get; set; }

    [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "directedgraph")]
    public MODELS.Layouts.DirectedGraph.DirectedGraphDocument DirectedGraphDocument { get; set; }

    [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "datatable")]
    public MODELS.Data.DataTableModel DataTableModel { get; set; }

    [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "systemxmldoc")]
    public MODELS.Data.XmlModel XmlModel;

    protected override void ProcessRecord()
    {
        var app = this.Client.Application.GetApplication();
        if (app == null)
        {
            string msg = "A Visio Application Instance is not attached";
            this.WriteVerbose(msg);
            throw new System.ArgumentOutOfRangeException(msg);
        }

        if (this.DirectedGraphDocument != null)
        {
            var dgstyline = new DirectedGraphStyling();
            this.Client.Model.DrawDirectedGraphDocument(this.DirectedGraphDocument, dgstyline);
        }
        else if (this.OrgChart != null)
        {
            this.Client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, this.OrgChart);
        }
        else if (this.GridLayout != null)
        {
            this.Client.Model.DrawGrid(VisioScripting.TargetPage.Auto, this.GridLayout);
        }
        else if (this.DataTableModel != null)
        {
            this.Client.Model.DrawDataTableModel(VisioScripting.TargetPage.Auto, this.DataTableModel);
        }
        else if (this.XmlModel != null)
        {
            this.Client.Model.DrawXmlModel(VisioScripting.TargetPage.Auto, this.XmlModel);
        }
        else
        {
            this.WriteVerbose("No object to draw");
        }
    }
}
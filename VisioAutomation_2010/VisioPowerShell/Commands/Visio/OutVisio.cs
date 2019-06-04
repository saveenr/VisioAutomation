using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsData.Out, Nouns.Visio)]
    public class OutVisio : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "orgchcart")]
        public VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument OrgChart { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "grid")]
        public VisioAutomation.Models.Layouts.Grid.GridLayout GridLayout { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "directedgraph")]
        public VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphDocument DirectedGraphDocument { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "datatable")]
        public VisioAutomation.Models.Data.DataTableModel DataTableModel { get; set; }

        [SMA.Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "systemxmldoc")]
        public VisioAutomation.Models.Data.XmlModel XmlModel;

        protected override void ProcessRecord()
        {
            var app = this.Client.Application.GetAttachedApplication();
            if (app == null)
            {
                string msg = "A Visio Application Instance is not attached";
                this.WriteVerbose(msg);
                throw new System.ArgumentOutOfRangeException(msg);
            }

            if (this.DirectedGraphDocument != null)
            {
                this.Client.Model.DrawDirectedGraphDocument(this.DirectedGraphDocument);
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
}
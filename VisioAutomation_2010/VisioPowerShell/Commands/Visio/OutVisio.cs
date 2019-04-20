using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Visio
{
    [SMA.Cmdlet(SMA.VerbsData.Out, Nouns.Visio)]
    public class OutVisio : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "orgchcart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Documents.OrgCharts.OrgChartDocument OrgChart { get; set; }

        [SMA.Parameter(ParameterSetName = "grid", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Layouts.Grid.GridLayout GridLayout { get; set; }

        [SMA.Parameter(ParameterSetName = "directedgraph", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public List<VisioAutomation.Models.Layouts.DirectedGraph.DirectedGraphLayout> DirectedGraphs { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioScripting.Models.DataTableModel DataTableModel { get; set; }

        [SMA.Parameter(ParameterSetName = "systemxmldoc", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioScripting.Models.XmlModel XmlModel;

        protected override void ProcessRecord()
        {
            var app = this.Client.Application.GetAttachedApplication();
            if (app == null)
            {
                string msg = "A Visio Application Instance is not attached";
                this.WriteVerbose(msg);
                throw new System.ArgumentOutOfRangeException(msg);
            }

            if (this.OrgChart != null)
            {
                this.Client.Model.DrawOrgChart(VisioScripting.TargetPage.Auto, this.OrgChart);
            }
            else if (this.GridLayout != null)
            {
                this.Client.Model.DrawGrid(VisioScripting.TargetPage.Auto, this.GridLayout);
            }
            else if (this.DirectedGraphs != null)
            {
                this.Client.Model.DrawDirectedGraphDocument(this.DirectedGraphs);
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
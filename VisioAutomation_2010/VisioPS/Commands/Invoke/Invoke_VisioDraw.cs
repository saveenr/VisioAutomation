using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioDraw")]
    public class Invoke_VisioDraw : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(ParameterSetName="orgchcart",Position=0,Mandatory = true)]
        public VA.Models.OrgChart.OrgChartDocument OrgChart { get; set; }

        [SMA.Parameter(ParameterSetName = "grid", Position = 0, Mandatory = true)]
        public VA.Models.Grid.GridLayout GridLayout { get; set; }

        [SMA.Parameter(ParameterSetName = "directedgraph", Position = 0, Mandatory = true)]
        public List<VA.Models.DirectedGraph.Drawing> DirectedGraphs { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 0, Mandatory = true)]
        public System.Data.DataTable DataTable { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 1, Mandatory = true)]
        public double CellWidth { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 2, Mandatory = true)]
        public double CellHeight { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 3, Mandatory = true)]
        public double CellSpacing { get; set; }

        [SMA.Parameter(ParameterSetName = "piechart", Position = 0, Mandatory = true)]
        public VA.Models.Charting.PieChart PieChart;

        [SMA.Parameter(ParameterSetName = "barchart", Position = 0, Mandatory = true)]
        public VA.Models.Charting.BarChart BarChart;

        [SMA.Parameter(ParameterSetName = "areachart", Position = 0, Mandatory = true)]
        public VA.Models.Charting.AreaChart AreaChart;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.OrgChart != null)
            {
                scriptingsession.Draw.OrgChart(this.OrgChart);                
            }
            else if (this.GridLayout != null)
            {
                scriptingsession.Draw.Grid(this.GridLayout);
            }
            else if (this.DirectedGraphs != null)
            {
                scriptingsession.Draw.DirectedGraph(this.DirectedGraphs);
            }
            else if (this.DataTable != null)
            {
                var widths = Enumerable.Repeat<double>(CellWidth, DataTable.Columns.Count).ToList();
                var heights = Enumerable.Repeat<double>(CellHeight, DataTable.Rows.Count).ToList();
                var spacing = new VA.Drawing.Size(CellSpacing, CellSpacing);
                var shapes = scriptingsession.Draw.Table(DataTable, widths, heights, spacing);
                this.WriteObject(shapes);
            }
            else if (this.PieChart != null)
            {
                scriptingsession.Draw.PieChart(this.PieChart);
            }
            else if (this.BarChart != null)
            {
                scriptingsession.Draw.BarChart(this.BarChart);
            }
            else if (this.AreaChart != null)
            {
                scriptingsession.Draw.AreaChart(this.AreaChart);
            }
            else
            {
                this.WriteVerboseEx("No object to draw");
            }
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Out, "Visio")]
    public class Out_Visio : VisioCmdlet
    {
        [SMA.Parameter(ParameterSetName = "orgchcart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VA.Models.OrgChart.OrgChartDocument OrgChart { get; set; }

        [SMA.Parameter(ParameterSetName = "grid", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VA.Models.Grid.GridLayout GridLayout { get; set; }

        [SMA.Parameter(ParameterSetName = "directedgraph", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public List<VA.Models.DirectedGraph.Drawing> DirectedGraphs { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public System.Data.DataTable DataTable { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public double CellWidth { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 2, Mandatory = true, ValueFromPipeline = true)]
        public double CellHeight { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 3, Mandatory = true, ValueFromPipeline = true)]
        public double CellSpacing { get; set; }

        [SMA.Parameter(ParameterSetName = "piechart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VA.Models.Charting.PieChart PieChart;

        [SMA.Parameter(ParameterSetName = "barchart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VA.Models.Charting.BarChart BarChart;

        [SMA.Parameter(ParameterSetName = "areachart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VA.Models.Charting.AreaChart AreaChart;

        [SMA.Parameter(ParameterSetName = "systemxmldoc", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public System.Xml.XmlDocument XmlDocument;

        protected override void ProcessRecord()
        {
            if (this.OrgChart != null)
            {
                this.client.Draw.OrgChart(this.OrgChart);
            }
            else if (this.GridLayout != null)
            {
                this.client.Draw.Grid(this.GridLayout);
            }
            else if (this.DirectedGraphs != null)
            {
                this.client.Draw.DirectedGraph(this.DirectedGraphs);
            }
            else if (this.DataTable != null)
            {
                var widths = Enumerable.Repeat<double>(CellWidth, DataTable.Columns.Count).ToList();
                var heights = Enumerable.Repeat<double>(CellHeight, DataTable.Rows.Count).ToList();
                var spacing = new VA.Drawing.Size(CellSpacing, CellSpacing);
                var shapes = this.client.Draw.Table(DataTable, widths, heights, spacing);
                this.WriteObject(shapes);
            }
            else if (this.PieChart != null)
            {
                this.client.Draw.PieChart(this.PieChart);
            }
            else if (this.BarChart != null)
            {
                this.client.Draw.BarChart(this.BarChart);
            }
            else if (this.AreaChart != null)
            {
                this.client.Draw.AreaChart(this.AreaChart);
            }
            else if (this.XmlDocument != null)
            {
                this.WriteVerbose("XmlDocument");
                var tree_drawing = new VA.Models.Tree.Drawing();
                build_from_xml_doc(this.XmlDocument, tree_drawing);

                tree_drawing.Render(this.client.Page.Get());
            }
            else
            {
                this.WriteVerbose("No object to draw");
            }
        }

        private void build_from_xml_doc(System.Xml.XmlDocument xmlDocument, VA.Models.Tree.Drawing tree_drawing)
        {
            var n = new VA.Models.Tree.Node();
            tree_drawing.Root = n;
            n.Text = new VA.Text.Markup.TextElement(xmlDocument.Name);
            this.build_from_xml_element(xmlDocument.DocumentElement,n);

        }

        private void build_from_xml_element(System.Xml.XmlElement x, VA.Models.Tree.Node parent)
        {
            foreach (System.Xml.XmlNode xchild in x.ChildNodes)
            {
                if (xchild is System.Xml.XmlElement)
                {
                    var nchild = new VA.Models.Tree.Node();
                    nchild.Text = new VA.Text.Markup.TextElement(xchild.Name);

                    parent.Children.Add(nchild);
                    build_from_xml_element( (System.Xml.XmlElement) xchild, nchild);
                }
            }
        }
    }
}
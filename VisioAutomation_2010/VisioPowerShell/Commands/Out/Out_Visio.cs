using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Management.Automation;
using System.Xml;

namespace VisioPowerShell.Commands.Out
{
    [Cmdlet(VerbsData.Out, "Visio")]
    public class Out_Visio : VisioCmdlet
    {
        [Parameter(ParameterSetName = "orgchcart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.OrgChart.OrgChartDocument OrgChart { get; set; }

        [Parameter(ParameterSetName = "grid", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Grid.GridLayout GridLayout { get; set; }

        [Parameter(ParameterSetName = "directedgraph", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public List<VisioAutomation.Models.DirectedGraph.Drawing> DirectedGraphs { get; set; }

        [Parameter(ParameterSetName = "datatable", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public DataTable DataTable { get; set; }

        [Parameter(ParameterSetName = "datatable", Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public double CellWidth { get; set; }

        [Parameter(ParameterSetName = "datatable", Position = 2, Mandatory = true, ValueFromPipeline = true)]
        public double CellHeight { get; set; }

        [Parameter(ParameterSetName = "datatable", Position = 3, Mandatory = true, ValueFromPipeline = true)]
        public double CellSpacing { get; set; }

        [Parameter(ParameterSetName = "piechart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Charting.PieChart PieChart;

        [Parameter(ParameterSetName = "barchart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Charting.BarChart BarChart;

        [Parameter(ParameterSetName = "areachart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public VisioAutomation.Models.Charting.AreaChart AreaChart;

        [Parameter(ParameterSetName = "systemxmldoc", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public XmlDocument XmlDocument;

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
                var widths = Enumerable.Repeat<double>(this.CellWidth, this.DataTable.Columns.Count).ToList();
                var heights = Enumerable.Repeat<double>(this.CellHeight, this.DataTable.Rows.Count).ToList();
                var spacing = new VisioAutomation.Drawing.Size(this.CellSpacing, this.CellSpacing);
                var shapes = this.client.Draw.Table(this.DataTable, widths, heights, spacing);
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
                var tree_drawing = new VisioAutomation.Models.Tree.Drawing();
                this.build_from_xml_doc(this.XmlDocument, tree_drawing);

                tree_drawing.Render(this.client.Page.Get());
            }
            else
            {
                this.WriteVerbose("No object to draw");
            }
        }

        private void build_from_xml_doc(XmlDocument xmlDocument, VisioAutomation.Models.Tree.Drawing tree_drawing)
        {
            var n = new VisioAutomation.Models.Tree.Node();
            tree_drawing.Root = n;
            n.Text = new VisioAutomation.Text.Markup.TextElement(xmlDocument.Name);
            this.build_from_xml_element(xmlDocument.DocumentElement,n);

        }

        private void build_from_xml_element(XmlElement x, VisioAutomation.Models.Tree.Node parent)
        {
            foreach (XmlNode xchild in x.ChildNodes)
            {
                if (xchild is XmlElement)
                {
                    var nchild = new VisioAutomation.Models.Tree.Node();
                    nchild.Text = new VisioAutomation.Text.Markup.TextElement(xchild.Name);

                    parent.Children.Add(nchild);
                    this.build_from_xml_element( (XmlElement) xchild, nchild);
                }
            }
        }
    }
}
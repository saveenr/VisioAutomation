using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
using VisioAutomation.Drawing;
using VisioAutomation.Models.Charting;
using VisioAutomation.Models.DirectedGraph;
using VisioAutomation.Models.Grid;
using VisioAutomation.Models.OrgChart;
using VisioAutomation.Text.Markup;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsData.Out, "Visio")]
    public class Out_Visio : VisioCmdlet
    {
        [SMA.ParameterAttribute(ParameterSetName = "orgchcart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public OrgChartDocument OrgChart { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "grid", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public GridLayout GridLayout { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "directedgraph", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public List<Drawing> DirectedGraphs { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "datatable", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public DataTable DataTable { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "datatable", Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public double CellWidth { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "datatable", Position = 2, Mandatory = true, ValueFromPipeline = true)]
        public double CellHeight { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "datatable", Position = 3, Mandatory = true, ValueFromPipeline = true)]
        public double CellSpacing { get; set; }

        [SMA.ParameterAttribute(ParameterSetName = "piechart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public PieChart PieChart;

        [SMA.ParameterAttribute(ParameterSetName = "barchart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public BarChart BarChart;

        [SMA.ParameterAttribute(ParameterSetName = "areachart", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public AreaChart AreaChart;

        [SMA.ParameterAttribute(ParameterSetName = "systemxmldoc", Position = 0, Mandatory = true, ValueFromPipeline = true)]
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
                var spacing = new Size(this.CellSpacing, this.CellSpacing);
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
            n.Text = new TextElement(xmlDocument.Name);
            this.build_from_xml_element(xmlDocument.DocumentElement,n);

        }

        private void build_from_xml_element(XmlElement x, VisioAutomation.Models.Tree.Node parent)
        {
            foreach (XmlNode xchild in x.ChildNodes)
            {
                if (xchild is XmlElement)
                {
                    var nchild = new VisioAutomation.Models.Tree.Node();
                    nchild.Text = new TextElement(xchild.Name);

                    parent.Children.Add(nchild);
                    this.build_from_xml_element( (XmlElement) xchild, nchild);
                }
            }
        }
    }
}
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
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
        public DataTable DataTable { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public double CellWidth { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 2, Mandatory = true, ValueFromPipeline = true)]
        public double CellHeight { get; set; }

        [SMA.Parameter(ParameterSetName = "datatable", Position = 3, Mandatory = true, ValueFromPipeline = true)]
        public double CellSpacing { get; set; }

        [SMA.Parameter(ParameterSetName = "systemxmldoc", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public XmlDocument XmlDocument;

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
                var targetpage = new VisioScripting.TargetPage();
                this.Client.Model.NewOrgChartDocument(targetpage, this.OrgChart);
            }
            else if (this.GridLayout != null)
            {
                var targetpage = new VisioScripting.TargetPage();
                this.Client.Model.DrawGrid(targetpage, this.GridLayout);
            }
            else if (this.DirectedGraphs != null)
            {
                this.Client.Model.NewDirectedGraphDocument(this.DirectedGraphs);
            }
            else if (this.DataTable != null)
            {
                var widths = Enumerable.Repeat<double>(this.CellWidth, this.DataTable.Columns.Count).ToList();
                var heights = Enumerable.Repeat<double>(this.CellHeight, this.DataTable.Rows.Count).ToList();
                var spacing = new VisioAutomation.Geometry.Size(this.CellSpacing, this.CellSpacing);
                var targetpage = new VisioScripting.TargetPage();
                var shapes = this.Client.Model.DrawDataTable(targetpage, this.DataTable, widths, heights, spacing);
                this.WriteObject(shapes);
            }
            else if (this.XmlDocument != null)
            {
                this.WriteVerbose("XmlDocument");
                var tree_drawing = new VisioAutomation.Models.Layouts.Tree.Drawing();
                this._build_from_xml_doc(this.XmlDocument, tree_drawing);

                tree_drawing.Render(this.Client.Page.GetActivePage());
            }
            else
            {
                this.WriteVerbose("No object to draw");
            }
        }

        private void _build_from_xml_doc(XmlDocument xml_document, VisioAutomation.Models.Layouts.Tree.Drawing tree_drawing)
        {
            var n = new VisioAutomation.Models.Layouts.Tree.Node();
            tree_drawing.Root = n;
            n.Text = new VisioAutomation.Models.Text.Element(xml_document.Name);
            this._build_from_xml_element(xml_document.DocumentElement,n);

        }

        private void _build_from_xml_element(XmlElement x, VisioAutomation.Models.Layouts.Tree.Node parent)
        {
            foreach (XmlNode xchild in x.ChildNodes)
            {
                if (xchild is XmlElement)
                {
                    var nchild = new VisioAutomation.Models.Layouts.Tree.Node();
                    nchild.Text = new VisioAutomation.Models.Text.Element(xchild.Name);

                    parent.Children.Add(nchild);
                    this._build_from_xml_element( (XmlElement) xchild, nchild);
                }
            }
        }
    }
}
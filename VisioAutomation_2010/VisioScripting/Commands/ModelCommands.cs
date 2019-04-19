using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using ORG = VisioAutomation.Models.Documents.OrgCharts;
using GRAPH = VisioAutomation.Models.Layouts.DirectedGraph;
using GRID = VisioAutomation.Models.Layouts.Grid;

namespace VisioScripting.Commands
{
    public class ModelCommands : CommandSet
    {
        internal ModelCommands(Client client) :
            base(client)
        {

        }

        public List<IVisio.Shape> DrawDataTable( VisioScripting.TargetPage targetpage, 
            System.Data.DataTable datatable,
            IList<double> widths,
            IList<double> heights,
            VisioAutomation.Geometry.Size cellspacing)
        {
            if (datatable == null)
            {
                throw new System.ArgumentNullException(nameof(datatable));
            }

            if (widths == null)
            {
                throw new System.ArgumentNullException(nameof(widths));
            }

            if (heights == null)
            {
                throw new System.ArgumentNullException(nameof(heights));
            }

            if (datatable.Rows.Count < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(datatable),"DataTable must have at least one row");
            }

            targetpage = targetpage.Resolve(this._client);
            string master = "Rectangle";
            string stencil = "basic_u.vss";
            var stencildoc = this._client.Document.OpenStencilDocument(stencil);
            var stencildoc_masters = stencildoc.Masters;
            var masterobj = stencildoc_masters.ItemU[master];
            
            targetpage.Page.Background = 0; // ensure this is a foreground page

            var pagesize = VisioAutomation.Pages.PageHelper.GetSize(targetpage.Page);

            var layout = new GRID.GridLayout(datatable.Columns.Count, datatable.Rows.Count, new VisioAutomation.Geometry.Size(1, 1), masterobj);
            layout.Origin = new VisioAutomation.Geometry.Point(0, pagesize.Height);
            layout.CellSpacing = cellspacing;
            layout.RowDirection = GRID.RowDirection.TopToBottom;
            layout.PerformLayout();

            foreach (var i in Enumerable.Range(0, datatable.Rows.Count))
            {
                var row = datatable.Rows[i];

                for (int col_index = 0; col_index < row.ItemArray.Length; col_index++)
                {
                    var col = row.ItemArray[col_index];
                    var cur_label = (col != null) ? col.ToString() : string.Empty;
                    var node = layout.GetNode(col_index, i);
                    node.Text = cur_label;
                }
            }

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawDataTable)))
            {
                layout.Render(targetpage.Page);
                targetpage.Page.ResizeToFitContents();
            }

            var shapes = layout.Nodes.Select(n => n.Shape).ToList();
            return shapes;
        }

        public void DrawGrid(VisioScripting.TargetPage targetpage, GRID.GridLayout layout)
        {
            targetpage = targetpage.Resolve(this._client);
            layout.PerformLayout();

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawGrid)))
            {
                layout.Render(targetpage.Page);
            }
        }

        public void DrawDataTableModel(VisioScripting.TargetPage targetpage, Models.DataTableModel dt_model)
        {
            targetpage = targetpage.Resolve(this._client);

            var widths = Enumerable.Repeat<double>(dt_model.CellWidth, dt_model.DataTable.Columns.Count).ToList();
            var heights = Enumerable.Repeat<double>(dt_model.CellHeight, dt_model.DataTable.Rows.Count).ToList();
            var spacing = new VisioAutomation.Geometry.Size(dt_model.CellSpacing, dt_model.CellSpacing);
            var shapes = this._client.Model.DrawDataTable(VisioScripting.TargetPage.Auto, dt_model.DataTable, widths, heights, spacing);
        }

        public void DrawXmlModel(VisioScripting.TargetPage targetpage, Models.XmlModel xmlmodel)
        {
            targetpage = targetpage.Resolve(this._client);

            var tree_drawing = new VisioAutomation.Models.Layouts.Tree.Drawing();
            this._build_from_xml_doc(xmlmodel.XmlDocument, tree_drawing);

            tree_drawing.Render(targetpage.Page);

        }

        private void _build_from_xml_doc(System.Xml.XmlDocument xml_document, VisioAutomation.Models.Layouts.Tree.Drawing tree_drawing)
        {
            var n = new VisioAutomation.Models.Layouts.Tree.Node();
            tree_drawing.Root = n;
            n.Text = new VisioAutomation.Models.Text.Element(xml_document.Name);
            this._build_from_xml_element(xml_document.DocumentElement, n);

        }

        private void _build_from_xml_element(System.Xml.XmlElement x, VisioAutomation.Models.Layouts.Tree.Node parent)
        {
            foreach (System.Xml.XmlNode xchild in x.ChildNodes)
            {
                if (xchild is System.Xml.XmlElement)
                {
                    var nchild = new VisioAutomation.Models.Layouts.Tree.Node();
                    nchild.Text = new VisioAutomation.Models.Text.Element(xchild.Name);

                    parent.Children.Add(nchild);
                    this._build_from_xml_element((System.Xml.XmlElement)xchild, nchild);
                }
            }
        }

        public void DrawOrgChart(VisioScripting.TargetPage targetpage, ORG.OrgChartDocument chartdocument)
        {
            targetpage = targetpage.Resolve(this._client);

            this._client.Output.WriteVerbose("Start OrgChart Rendering");

            var application = targetpage.Page.Application;
            chartdocument.Render(application);

            targetpage.Page.ResizeToFitContents();
            this._client.Output.WriteVerbose("Finished OrgChart Rendering");
        }

        public void NewDirectedGraphDocument(IList<GRAPH.DirectedGraphLayout> graph)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            this._client.Output.WriteVerbose("Start rendering directed graph");
            var app = cmdtarget.Application;

            this._client.Output.WriteVerbose("Creating a New Document For the Directed Graphs");
            string template = null;
            var doc = this._client.Document.NewDocumentFromTemplate(template);

            int num_pages_created = 0;
            var doc_pages = doc.Pages;

            foreach (int i in Enumerable.Range(0, graph.Count))
            {
                var dg = graph[i];

                
                var options = new GRAPH.MsaglLayoutOptions();
                options.UseDynamicConnectors = false;

                // if this is the first page to drawe
                // then reuse the initial empty page in the document
                // otherwise, create a new page.
                var page = num_pages_created == 0 ? app.ActivePage : doc_pages.Add();

                this._client.Output.WriteVerbose("Rendering page: {0}", i + 1);
                dg.Render(page, options);

                var targetpages = new VisioScripting.TargetPages(page);
                this._client.Page.ResizePageToFitContents(targetpages, new VisioAutomation.Geometry.Size(1.0, 1.0));
                this._client.View.SetZoomToObject(VisioScripting.TargetWindow.Auto, VisioScripting.Models.ZoomToObject.Page);
                this._client.Output.WriteVerbose("Finished rendering page");

                num_pages_created++;
            }

            this._client.Output.WriteVerbose("Finished rendering all pages");
            this._client.Output.WriteVerbose("Finished rendering directed graph.");
        }
    }
}
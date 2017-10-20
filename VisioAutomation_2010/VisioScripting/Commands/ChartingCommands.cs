using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using ORG = VisioAutomation.Models.Documents.OrgCharts;
using GRAPH = VisioAutomation.Models.Layouts.DirectedGraph;
using GRID = VisioAutomation.Models.Layouts.Grid;

namespace VisioScripting.Commands
{
    public class ChartingCommands : CommandSet
    {
        internal ChartingCommands(Client client) :
            base(client)
        {

        }

        public List<IVisio.Shape> NewDataTablePageInActiveDocument(
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

            var cmdtarget = this._client.GetCommandTargetPage();
            string master = "Rectangle";
            string stencil = "basic_u.vss";
            var stencildoc = this._client.Document.OpenStencilDocument(stencil);
            var stencildoc_masters = stencildoc.Masters;
            var masterobj = stencildoc_masters.ItemU[master];

            var active_document = cmdtarget.ActiveDocument;
            var pages = active_document.Pages;

            var page = pages.Add();
            page.Background = 0; // ensure this is a foreground page

            var pagesize = this._client.Page.GetActivePageSize();

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

            using (var undoscope = this._client.Application.NewUndoScope(nameof(NewDataTablePageInActiveDocument)))
            {
                layout.Render(page);
                page.ResizeToFitContents();
            }

            var shapes = layout.Nodes.Select(n => n.Shape).ToList();
            return shapes;
        }

        public void DrawGridOnActivePage(GRID.GridLayout layout)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            layout.PerformLayout();

            using (var undoscope = this._client.Application.NewUndoScope(nameof(DrawGridOnActivePage)))
            {
                layout.Render(page);
            }
        }

        public IVisio.Shape DrawPieSliceOnActivePage(
            VisioAutomation.Geometry.Point center,
            double radius,
            double start_angle,
            double end_angle)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var application = cmdtarget.Application;
            using (var undoscope = this._client.Application.NewUndoScope(nameof(DrawPieSliceOnActivePage)))
            {
                var active_page = application.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }
        public IVisio.Shape DrawDoughnutSliceOnActivePage(
            VisioAutomation.Geometry.Point center,
            double inner_radius,
            double outer_radius,
            double start_angle,
            double end_angle)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var application = cmdtarget.Application;
            using (var undoscope = this._client.Application.NewUndoScope(nameof(DrawDoughnutSliceOnActivePage)))
            {
                var active_page = cmdtarget.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, inner_radius, outer_radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }

        public void DrawPieChartOnActivePage(VisioAutomation.Models.Charting.PieChart chart)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            chart.Render(page);
        }

        public void DrawBarChartOnActivePage(VisioAutomation.Models.Charting.BarChart chart)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var application = cmdtarget.Application;
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void DrawAreaChartOnActivePage(VisioAutomation.Models.Charting.AreaChart chart)
        {
            var cmdtarget = this._client.GetCommandTargetPage();

            var page = cmdtarget.ActivePage;
            chart.Render(page);
        }


        public void NewOrgChartDocument(ORG.OrgChartDocument chartdocument)
        {
            var cmdtarget = this._client.GetCommandTargetApplication();

            this._client.Output.WriteVerbose("Start OrgChart Rendering");

            var application = cmdtarget.Application;
            chartdocument.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
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
                this._client.Page.ResizeActivePageToFitContents(new VisioAutomation.Geometry.Size(1.0, 1.0), true);
                this._client.View.Zoom(VisioScripting.Models.Zoom.ToPage);
                this._client.Output.WriteVerbose("Finished rendering page");

                num_pages_created++;
            }

            this._client.Output.WriteVerbose("Finished rendering all pages");
            this._client.Output.WriteVerbose("Finished rendering directed graph.");
        }
    }
}
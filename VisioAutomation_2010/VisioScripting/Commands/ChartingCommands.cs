using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
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

        public VisioAutomation.SurfaceTarget GetActiveDrawingSurface()
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            var surf_Application = cmdtarget.Application;
            var surf_Window = surf_Application.ActiveWindow;
            var surf_Window_subtype = surf_Window.SubType;

            // TODO: Revisit the logic here
            // TODO: And what about a selected shape as a surface?

            this._client.Output.WriteVerbose("Window SubType: {0}", surf_Window_subtype);
            if (surf_Window_subtype == 64)
            {
                this._client.Output.WriteVerbose("Window = Master Editing");
                var surf_Master = (IVisio.Master)surf_Window.Master;
                var surface = new VisioAutomation.SurfaceTarget(surf_Master);
                return surface;

            }
            else
            {
                this._client.Output.WriteVerbose("Window = Page ");
                var surf_Page = surf_Application.ActivePage;
                var surface = new VisioAutomation.SurfaceTarget(surf_Page);
                return surface;
            }
        }

        public List<IVisio.Shape> DrawTable(System.Data.DataTable datatable,
                                          IList<double> widths,
                                          IList<double> heights,
            VisioAutomation.Geometry.Size cellspacing)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

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
                return new List<IVisio.Shape>(0);
            }


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

            using (var undoscope = this._client.Application.NewUndoScope("Draw Table"))
            {
                layout.Render(page);
                page.ResizeToFitContents();
            }

            var page_shapes = page.Shapes;
            var shapes = layout.Nodes.Select(n => n.Shape).ToList();
            return shapes;

        }

        public void DrawGrid(GRID.GridLayout layout)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            //Create a new page to hold the grid
            var page = cmdtarget.ActivePage;
            layout.PerformLayout();

            using (var undoscope = this._client.Application.NewUndoScope("Draw Grid"))
            {
                layout.Render(page);
            }
        }

        public IVisio.Shape DrawPieSlice(VisioAutomation.Geometry.Point center,
                                  double radius,
                                  double start_angle,
                                  double end_angle)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var application = cmdtarget.Application;
            using (var undoscope = this._client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }
        public IVisio.Shape DrawDoughnutSlice(VisioAutomation.Geometry.Point center,
                          double inner_radius,
                          double outer_radius,
                          double start_angle,
                          double end_angle)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);
            
            var application = cmdtarget.Application;
            using (var undoscope = this._client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = cmdtarget.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, inner_radius, outer_radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }

        public void DrawPieChart(VisioAutomation.Models.Charting.PieChart chart)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var page = cmdtarget.ActivePage;
            chart.Render(page);
        }

        public void DrawBarChart(VisioAutomation.Models.Charting.BarChart chart)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var application = cmdtarget.Application;
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void DrawAreaChart(VisioAutomation.Models.Charting.AreaChart chart)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);

            var page = cmdtarget.ActivePage;
            chart.Render(page);
        }


        public void DrawOrgChart(ORG.OrgChartDocument orgChartDocument)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

            this._client.Output.WriteVerbose("Start OrgChart Rendering");

            var application = cmdtarget.Application;
            orgChartDocument.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
            this._client.Output.WriteVerbose("Finished OrgChart Rendering");
        }

        public void DrawDirectedGraph(IList<GRAPH.DirectedGraphLayout> graph)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application);

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
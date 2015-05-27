using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using GRIDLAYOUT = VisioAutomation.Models.Grid;
using ORGCHARTLAYOUT = VisioAutomation.Models.OrgChart;
using DGMODEL = VisioAutomation.Models.DirectedGraph;

namespace VisioAutomation.Scripting.Commands
{
    public class DrawCommands : CommandSet
    {
        internal DrawCommands(Client client) :
            base(client)
        {

        }

        public Drawing.DrawingSurface GetDrawingSurface()
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var surf_Application = this.Client.Application.Get();
            var surf_Window = surf_Application.ActiveWindow;
            var surf_Window_subtype = surf_Window.SubType;

            // TODO: Revisit the logic here
            // TODO: And what about a selected shape as a surface?

            this.Client.WriteVerbose("Window SubType: {0}", surf_Window_subtype);
            if (surf_Window_subtype == 64)
            {
                this.Client.WriteVerbose("Window = Master Editing");
                var surf_Master = (IVisio.Master)surf_Window.Master;
                var surface = new Drawing.DrawingSurface(surf_Master);
                return surface;

            }
            else
            {
                this.Client.WriteVerbose("Window = Page ");
                var surf_Page = surf_Application.ActivePage;
                var surface = new Drawing.DrawingSurface(surf_Page);
                return surface;
            }
        }

        public IList<IVisio.Shape> Table(System.Data.DataTable datatable,
                                          IList<double> widths,
                                          IList<double> heights,
                                          Drawing.Size cellspacing)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

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
            var stencildoc = this.Client.Document.OpenStencil(stencil);
            var stencildoc_masters = stencildoc.Masters;
            var masterobj = stencildoc_masters.ItemU[master];

            var app = this.Client.Application.Get();
            var application = app;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;

            var page = pages.Add();
            page.Background = 0; // ensure this is a foreground page

            var pagesize = this.Client.Page.GetSize();

            var layout = new GRIDLAYOUT.GridLayout(datatable.Columns.Count, datatable.Rows.Count, new Drawing.Size(1, 1), masterobj);
            layout.Origin = new Drawing.Point(0, pagesize.Height);
            layout.CellSpacing = cellspacing;
            layout.RowDirection = GRIDLAYOUT.RowDirection.TopToBottom;
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

            using (var undoscope = this.Client.Application.NewUndoScope("Draw Table"))
            {
                layout.Render(page);
                page.ResizeToFitContents();
            }

            var page_shapes = page.Shapes;
            var shapes = layout.Nodes.Select(n => n.Shape).ToList();
            return shapes;

        }

        public void Grid(GRIDLAYOUT.GridLayout layout)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            //Create a new page to hold the grid
            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            layout.PerformLayout();

            using (var undoscope = this.Client.Application.NewUndoScope("Draw Grid"))
            {
                layout.Render(page);
            }
        }

        public IVisio.Shape NURBSCurve(IList<Drawing.Point> controlpoints,
                                    IList<double> knots,
                                    IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw NURBS Curve"))
            {

                var page = application.ActivePage;
                var shape = page.DrawNURBS(controlpoints, knots, weights, degree);
                return shape;
            }
        }

        public IVisio.Shape Rectangle(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Rectangle"))
            {
                var shape = surface.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape Line(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Line"))
            {
                var shape = surface.DrawLine(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape Oval(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            var rect = new Drawing.Rectangle(x0, y0, x1, y1);
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Oval"))
            {
                var shape = surface.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape Oval(Drawing.Point center, double radius)
        {
            var surface = this.GetDrawingSurface();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Oval"))
            {
                var shape = surface.DrawOval(center, radius);
                return shape;
            }
        }

        public IVisio.Shape Bezier(IEnumerable<Drawing.Point> points)
        {
            var surface = this.GetDrawingSurface();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Bezier"))
            {
                var shape = surface.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape PolyLine(IList<Drawing.Point> points)
        {
            var surface = this.GetDrawingSurface();
            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw PolyLine"))
            {
                var shape = surface.DrawPolyLine(points);
                return shape;
            }
        }

        public IVisio.Shape PieSlice(Drawing.Point center,
                                  double radius,
                                  double start_angle,
                                  double end_angle)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }
        public IVisio.Shape DoughnutSlice(Drawing.Point center,
                          double inner_radius,
                          double outer_radius,
                          double start_angle,
                          double end_angle)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new Models.Charting.PieSlice(center, inner_radius, outer_radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }

        public void PieChart(Models.Charting.PieChart chart)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void BarChart(Models.Charting.BarChart chart)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void AreaChart(Models.Charting.AreaChart chart)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var application = this.Client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }


        public void OrgChart(ORGCHARTLAYOUT.OrgChartDocument orgChartDocument)
        {

            this.Client.WriteVerbose("Start OrgChart Rendering");
            this.Client.Application.AssertApplicationAvailable();

            var application = this.Client.Application.Get();
            orgChartDocument.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
            this.Client.WriteVerbose("Finished OrgChart Rendering");
        }

        public void DirectedGraph(IList<DGMODEL.Drawing> directedgraphs)
        {
            this.Client.Application.AssertApplicationAvailable();

            this.Client.WriteVerbose("Start rendering directed graph");
            var app = this.Client.Application.Get();


            this.Client.WriteVerbose("Creating a New Document For the Directed Graphs");
            var doc = this.Client.Document.New(null);

            int num_pages_created = 0;
            var doc_pages = doc.Pages;

            foreach (int i in Enumerable.Range(0, directedgraphs.Count))
            {
                var dg = directedgraphs[i];

                
                var options = new DGMODEL.MsaglLayoutOptions();
                options.UseDynamicConnectors = false;

                // if this is the first page to drawe
                // then reuse the initial empty page in the document
                // otherwise, create a new page.
                var page = num_pages_created == 0 ? app.ActivePage : doc_pages.Add();

                this.Client.WriteVerbose("Rendering page: {0}", i + 1);
                dg.Render(page, options);
                this.Client.Page.ResizeToFitContents(new Drawing.Size(1.0, 1.0), true);
                this.Client.View.Zoom(Zoom.ToPage);
                this.Client.WriteVerbose("Finished rendering page");

                num_pages_created++;
            }

            this.Client.WriteVerbose("Finished rendering all pages");
            this.Client.WriteVerbose("Finished rendering directed graph.");
        }

        public void Duplicate(int n)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(n));
            }
            if (!this.Client.Selection.HasShapes())
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = this.Client.Application.Get();
            using (var undoscope = this.Client.Application.NewUndoScope($"Duplicate Shape {n} Times"))
            {
                var active_window = application.ActiveWindow;
                var selection = active_window.Selection;
                var active_page = application.ActivePage;
                var new_shapes = DrawCommands.CreateDuplicates(active_page, selection[1], n);
            }
        }

        private static IList<IVisio.Shape> CreateDuplicates(IVisio.Page page,
                                           IVisio.Shape shape,
                                           int n)
        {
            // NOTE: n is the total number you want INCLUDING the original shape
            // example n=0 then result={s0}
            // example n=1, result={s0}
            // example n=2, result={s0,s1}
            // example n=3, result={s0,s1,s2}
            // where s0 is the original shape

            if (n < 2)
            {
                return new List<IVisio.Shape> { shape };
            }

            int num_doubles = (int)System.Math.Log(n, 2.0);
            int leftover = n - ((int)System.Math.Pow(2.0, num_doubles));
            if (leftover < 0)
            {
                throw new System.InvalidOperationException("internal error: leftover value must greater than or equal to zero");
            }

            var duplicated_shapes = new List<IVisio.Shape> { shape };

            var application = page.Application;
            var win = application.ActiveWindow;

            foreach (int i in Enumerable.Range(0, num_doubles))
            {
                win.DeselectAll();
                win.Select(duplicated_shapes, IVisio.VisSelectArgs.visSelect);
                var selection = win.Selection;
                selection.Duplicate();
                var selection1 = win.Selection;
                duplicated_shapes.AddRange(selection1.AsEnumerable());
            }

            if (leftover > 0)
            {
                var leftover_shapes = duplicated_shapes.Take(leftover);
                win.DeselectAll();
                win.Select(leftover_shapes, IVisio.VisSelectArgs.visSelect);
                var selection = win.Selection;
                selection.Duplicate();
                var selection1 = win.Selection;
                duplicated_shapes.AddRange(selection1.AsEnumerable());
            }

            win.DeselectAll();
            win.Select(duplicated_shapes, IVisio.VisSelectArgs.visSelect);

            if (duplicated_shapes.Count != n)
            {
                string msg = $"internal error: failed to create {n} shapes, instead created {duplicated_shapes.Count}";
                throw new VisioOperationException(msg);
            }

            var selection2 = win.Selection;
            if (selection2.Count != n)
            {
                throw new VisioOperationException("internal error: failed to select the duplicated shapes");
            }

            return duplicated_shapes;
        }

        public List<IVisio.Shape> GetAllShapes()
        {
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            return surface.GetAllShapes();
        }
    }
}
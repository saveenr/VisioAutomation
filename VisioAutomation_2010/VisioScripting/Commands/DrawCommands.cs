using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using ORGCHART = VisioAutomation.Models.Documents.OrgCharts;
using DG = VisioAutomation.Models.Layouts.DirectedGraph;
using GRID = VisioAutomation.Models.Layouts.Grid;

namespace VisioScripting.Commands
{
    public class DrawCommands : CommandSet
    {
        internal DrawCommands(Client client) :
            base(client)
        {

        }

        public VisioAutomation.SurfaceTarget GetDrawingSurface()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var surf_Application = this._client.Application.Get();
            var surf_Window = surf_Application.ActiveWindow;
            var surf_Window_subtype = surf_Window.SubType;

            // TODO: Revisit the logic here
            // TODO: And what about a selected shape as a surface?

            this._client.WriteVerbose("Window SubType: {0}", surf_Window_subtype);
            if (surf_Window_subtype == 64)
            {
                this._client.WriteVerbose("Window = Master Editing");
                var surf_Master = (IVisio.Master)surf_Window.Master;
                var surface = new VisioAutomation.SurfaceTarget(surf_Master);
                return surface;

            }
            else
            {
                this._client.WriteVerbose("Window = Page ");
                var surf_Page = surf_Application.ActivePage;
                var surface = new VisioAutomation.SurfaceTarget(surf_Page);
                return surface;
            }
        }

        public List<IVisio.Shape> Table(System.Data.DataTable datatable,
                                          IList<double> widths,
                                          IList<double> heights,
            VisioAutomation.Geometry.Size cellspacing)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

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
            var stencildoc = this._client.Document.OpenStencil(stencil);
            var stencildoc_masters = stencildoc.Masters;
            var masterobj = stencildoc_masters.ItemU[master];

            var app = this._client.Application.Get();
            var application = app;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;

            var page = pages.Add();
            page.Background = 0; // ensure this is a foreground page

            var pagesize = this._client.Page.GetSize();

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

        public void Grid(GRID.GridLayout layout)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            //Create a new page to hold the grid
            var application = this._client.Application.Get();
            var page = application.ActivePage;
            layout.PerformLayout();

            using (var undoscope = this._client.Application.NewUndoScope("Draw Grid"))
            {
                layout.Render(page);
            }
        }

        public IVisio.Shape NURBSCurve(IList<VisioAutomation.Geometry.Point> controlpoints,
                                    IList<double> knots,
                                    IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Draw NURBS Curve"))
            {

                var page = application.ActivePage;
                var shape = page.DrawNURBS(controlpoints, knots, weights, degree);
                return shape;
            }
        }

        public IVisio.Shape Rectangle(VisioAutomation.Geometry.Rectangle r)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Rectangle"))
            {
                var shape = surface.DrawRectangle(r.Left, r.Bottom, r.Right, r.Top);
                return shape;
            }
        }

        public IVisio.Shape Rectangle(double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.Rectangle(rect);
        }

        public IVisio.Shape Line(double x0, double y0, double x1, double y1)
        {
            var p0 = new VisioAutomation.Geometry.Point(x0, y0);
            var p1 = new VisioAutomation.Geometry.Point(x1, y1);
            return this.Line(p0, p1);
        }

        public IVisio.Shape Line(VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Line"))
            {
                var shape = surface.DrawLine(p0,p1);
                return shape;
            }
        }

        public IVisio.Shape Oval(VisioAutomation.Geometry.Rectangle rect)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Oval"))
            {
                var shape = surface.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape Oval(double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.Oval(rect);
        }

        public IVisio.Shape Oval(VisioAutomation.Geometry.Point center, double radius)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Oval"))
            {
                var shape = surface.DrawOval(center, radius);
                return shape;
            }
        }

        public IVisio.Shape Bezier(IEnumerable<VisioAutomation.Geometry.Point> points)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Bezier"))
            {
                var shape = surface.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape PolyLine(IList<VisioAutomation.Geometry.Point> points)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = this._client.Application.NewUndoScope("Draw PolyLine"))
            {
                var shape = surface.DrawPolyLine(points);
                return shape;
            }
        }

        public IVisio.Shape PieSlice(VisioAutomation.Geometry.Point center,
                                  double radius,
                                  double start_angle,
                                  double end_angle)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }
        public IVisio.Shape DoughnutSlice(VisioAutomation.Geometry.Point center,
                          double inner_radius,
                          double outer_radius,
                          double start_angle,
                          double end_angle)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new VisioAutomation.Models.Charting.PieSlice(center, inner_radius, outer_radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }

        public void PieChart(VisioAutomation.Models.Charting.PieChart chart)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void BarChart(VisioAutomation.Models.Charting.BarChart chart)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void AreaChart(VisioAutomation.Models.Charting.AreaChart chart)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var application = this._client.Application.Get();
            var page = application.ActivePage;
            chart.Render(page);
        }


        public void OrgChart(ORGCHART.OrgChartDocument orgChartDocument)
        {

            this._client.WriteVerbose("Start OrgChart Rendering");
            this._client.Application.AssertApplicationAvailable();

            var application = this._client.Application.Get();
            orgChartDocument.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
            this._client.WriteVerbose("Finished OrgChart Rendering");
        }

        public void DirectedGraph(IList<DG.DirectedGraphLayout> graph)
        {
            this._client.Application.AssertApplicationAvailable();

            this._client.WriteVerbose("Start rendering directed graph");
            var app = this._client.Application.Get();


            this._client.WriteVerbose("Creating a New Document For the Directed Graphs");
            var doc = this._client.Document.New(null);

            int num_pages_created = 0;
            var doc_pages = doc.Pages;

            foreach (int i in Enumerable.Range(0, graph.Count))
            {
                var dg = graph[i];

                
                var options = new DG.MsaglLayoutOptions();
                options.UseDynamicConnectors = false;

                // if this is the first page to drawe
                // then reuse the initial empty page in the document
                // otherwise, create a new page.
                var page = num_pages_created == 0 ? app.ActivePage : doc_pages.Add();

                this._client.WriteVerbose("Rendering page: {0}", i + 1);
                dg.Render(page, options);
                this._client.Page.ResizeToFitContents(new VisioAutomation.Geometry.Size(1.0, 1.0), true);
                this._client.View.Zoom(VisioScripting.Models.Zoom.ToPage);
                this._client.WriteVerbose("Finished rendering page");

                num_pages_created++;
            }

            this._client.WriteVerbose("Finished rendering all pages");
            this._client.WriteVerbose("Finished rendering directed graph.");
        }

        public void Duplicate(int n)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(n));
            }
            if (!this._client.Selection.HasShapes())
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope(string.Format("Duplicate Shape {0} Times", n)))
            {
                var active_window = application.ActiveWindow;
                var selection = active_window.Selection;
                var active_page = application.ActivePage;
                var new_shapes = DrawCommands.CreateDuplicates(active_page, selection[1], n);
            }
        }

        private static List<IVisio.Shape> CreateDuplicates(IVisio.Page page,
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
                duplicated_shapes.AddRange(selection1.ToEnumerable());
            }

            if (leftover > 0)
            {
                var leftover_shapes = duplicated_shapes.Take(leftover);
                win.DeselectAll();
                win.Select(leftover_shapes, IVisio.VisSelectArgs.visSelect);
                var selection = win.Selection;
                selection.Duplicate();
                var selection1 = win.Selection;
                duplicated_shapes.AddRange(selection1.ToEnumerable());
            }

            win.DeselectAll();
            win.Select(duplicated_shapes, IVisio.VisSelectArgs.visSelect);

            if (duplicated_shapes.Count != n)
            {
                string msg = string.Format("internal error: failed to create {0} shapes, instead created {1}", n,
                    duplicated_shapes.Count);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }

            var selection2 = win.Selection;
            if (selection2.Count != n)
            {
                throw new VisioAutomation.Exceptions.VisioOperationException("internal error: failed to select the duplicated shapes");
            }

            return duplicated_shapes;
        }

        public List<IVisio.Shape> GetAllShapes()
        {
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var shapes = surface.Shapes;
            var list = new List<IVisio.Shape>();
            list.AddRange(shapes.ToEnumerable());
            return list;
        }
    }
}
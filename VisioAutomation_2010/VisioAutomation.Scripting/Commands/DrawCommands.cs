using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using GRIDLAYOUT = VisioAutomation.Models.Grid;
using RADIALLAYOUT = VisioAutomation.Models.Charting;
using ORGCHARTLAYOUT = VisioAutomation.Models.OrgChart;
using DGMODEL = VisioAutomation.Models.DirectedGraph;

namespace VisioAutomation.Scripting.Commands
{
    public class DrawCommands : CommandSet
    {
        public DrawCommands(Client client) :
            base(client)
        {

        }

        public IList<IVisio.Shape> Table(System.Data.DataTable datatable,
                                          IList<double> widths,
                                          IList<double> heights,
                                          VA.Drawing.Size cellspacing)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            if (datatable == null)
            {
                throw new System.ArgumentNullException("datatable");
            }

            if (widths == null)
            {
                throw new System.ArgumentNullException("widths");
            }

            if (heights == null)
            {
                throw new System.ArgumentNullException("heights");
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

            var application = this.Client.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;

            var page = pages.Add();
            page.Background = 0; // ensure this is a foreground page

            var pagesize = this.Client.Page.GetSize();

            var layout = new GRIDLAYOUT.GridLayout(datatable.Columns.Count, datatable.Rows.Count, new VA.Drawing.Size(1, 1), masterobj);
            layout.Origin = new VA.Drawing.Point(0, pagesize.Height);
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

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Table"))
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
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            //Create a new page to hold the grid
            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            layout.PerformLayout();

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Grid"))
            {
                layout.Render(page);
            }
        }

        public IVisio.Shape NURBSCurve(IList<VA.Drawing.Point> controlpoints,
                                    IList<double> knots,
                                    IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw NURBS Curve"))
            {

                var page = application.ActivePage;
                var shape = page.DrawNURBS(controlpoints, knots, weights, degree);
                return shape;
            }
        }

        public IVisio.Shape Rectangle(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Rectangle"))
            {
                var shape = surface.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape Line(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Line"))
            {
                var shape = surface.DrawLine(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape Oval(double x0, double y0, double x1, double y1)
        {
            var surface = this.GetDrawingSurface();
            var rect = new VA.Drawing.Rectangle(x0, y0, x1, y1);
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Oval"))
            {
                var shape = surface.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape Oval(VA.Drawing.Point center, double radius)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Oval"))
            {
                var shape = surface.DrawOval(center, radius);
                return shape;
            }
        }

        public IVisio.Shape Bezier(IEnumerable<VA.Drawing.Point> points)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Bezier"))
            {
                var shape = surface.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape PolyLine(IList<VA.Drawing.Point> points)
        {
            var surface = this.GetDrawingSurface();
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw PolyLine"))
            {
                var shape = surface.DrawPolyLine(points);
                return shape;
            }
        }

        public IVisio.Shape PieSlice(VA.Drawing.Point center,
                                  double radius,
                                  double start_angle,
                                  double end_angle)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new VA.Models.Charting.PieSlice(center, radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }
        public IVisio.Shape DoughnutSlice(VA.Drawing.Point center,
                          double inner_radius,
                          double outer_radius,
                          double start_angle,
                          double end_angle)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Draw Pie Slice"))
            {
                var active_page = application.ActivePage;
                var slice = new VA.Models.Charting.PieSlice(center, inner_radius, outer_radius, start_angle, end_angle);
                var shape = slice.Render(active_page);
                return shape;
            }
        }

        public void PieChart(VA.Models.Charting.PieChart chart)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void BarChart(VA.Models.Charting.BarChart chart)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            chart.Render(page);
        }

        public void AreaChart(VA.Models.Charting.AreaChart chart)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var application = this.Client.VisioApplication;
            var page = application.ActivePage;
            chart.Render(page);
        }


        public void OrgChart(ORGCHARTLAYOUT.OrgChartDocument orgChartDocument)
        {

            this.Client.WriteVerbose("Start OrgChart Rendering");
            this.AssertApplicationAvailable();

            var application = this.Client.VisioApplication;
            orgChartDocument.Render(application);
            var active_page = application.ActivePage;
            active_page.ResizeToFitContents();
            this.Client.WriteVerbose("Finished OrgChart Rendering");
        }

        public void DirectedGraph(IList<DGMODEL.Drawing> directedgraphs)
        {
            this.AssertApplicationAvailable();

            this.Client.WriteVerbose("Start rendering directed graph");
            var app = this.Client.VisioApplication;


            this.Client.WriteVerbose("Creating a New Document For the Directed Graphs");
            var doc = this.Client.Document.New(null);

            int num_pages_created = 0;
            var doc_pages = doc.Pages;

            foreach (int i in Enumerable.Range(0, directedgraphs.Count))
            {
                var dg = directedgraphs[i];

                
                var options = new DGMODEL.MSAGLLayoutOptions();
                options.UseDynamicConnectors = false;

                // if this is the first page to drawe
                // then reuse the initial empty page in the document
                // otherwise, create a new page.
                var page = num_pages_created == 0 ? app.ActivePage : doc_pages.Add();

                this.Client.WriteVerbose("Rendering page: {0}", i + 1);
                dg.Render(page, options);
                this.Client.Page.ResizeToFitContents(new VA.Drawing.Size(1.0, 1.0), true);
                this.Client.View.Zoom(VA.Scripting.Zoom.ToPage);
                this.Client.WriteVerbose("Finished rendering page");

                num_pages_created++;
            }

            this.Client.WriteVerbose("Finished rendering all pages");
            this.Client.WriteVerbose("Finished rendering directed graph.");
        }

        public void Duplicate(int n)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException("n");
            }
            if (!this.Client.HasSelectedShapes())
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, string.Format("Duplicate Shape {0} Times", n)))
            {
                var active_window = application.ActiveWindow;
                var selection = active_window.Selection;
                var active_page = application.ActivePage;
                DrawCommands.CreateDuplicates(active_page, selection[1], n);
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
                string msg = string.Format("internal error: failed to create {0} shapes, instead created {1}", n,
                                           duplicated_shapes.Count);
                throw new VA.Scripting.ScriptingException(msg);
            }

            var selection2 = win.Selection;
            if (selection2.Count != n)
            {
                throw new VA.Scripting.ScriptingException("internal error: failed to select the duplicated shapes");
            }

            return duplicated_shapes;
        }

        public List<IVisio.Shape> GetAllShapes()
        {
            var surface = this.GetDrawingSurface();
            return surface.GetAllShapes();
        }
    }
}
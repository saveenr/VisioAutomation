using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;


namespace VisioAutomation.Scripting.Commands
{
    public class DrawCommands : SessionCommands
    {
        public DrawCommands(Session session) :
            base(session)
        {

        }

        public IList<IVisio.Shape> DrawDataTable(DataTable datatable,
                                          IList<double> widths,
                                          IList<double> heights,
                                          VA.Drawing.Size cellspacing)
        {
            if (datatable == null)
            {
                throw new ArgumentNullException("datatable");
            }

            if (widths == null)
            {
                throw new ArgumentNullException("widths");
            }

            if (heights == null)
            {
                throw new ArgumentNullException("heights");
            }

            if (datatable.Rows.Count < 1)
            {
                return new List<IVisio.Shape>(0);
            }
            
            string master = "Rectangle";
            string stencil = "basic_u.vss";
            var stencildoc = this.Session.Document.OpenStencil(stencil);
            var stencildoc_masters = stencildoc.Masters;
            var masterobj = stencildoc_masters.ItemU[master];

            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var pages = active_document.Pages;

            var page = pages.Add();
            page.Background = 0; // ensure this is a foreground page

            var pagesize = page.GetSize();
            
            var origin = new VA.Drawing.Point(0, pagesize.Height);

            var layout = new VA.Layout.Grid.GridLayout(datatable.Columns.Count, datatable.Rows.Count, new VA.Drawing.Size(1,1), masterobj);
            layout.PerformLayout(origin, cellspacing);
            layout.RowDirection = VA.Layout.Grid.RowDirection.TopToBottom;

            foreach (var i in Enumerable.Range(0, datatable.Rows.Count))
            {
                var row = datatable.Rows[i];

                for (int col_index = 0; col_index < row.ItemArray.Length; col_index++)
                {
                    var col = row.ItemArray[col_index];
                    var cur_label = (col != null) ? col.ToString() : String.Empty;
                    var node = layout.GetNode(col_index, i);
                    node.Text = cur_label;
                }
            }

            using (var undoscope = application.CreateUndoScope())
            {
                layout.Render(page);
                page.ResizeToFitContents(new VA.Drawing.Size(0,0));
            }

            var page_shapes = page.Shapes;
            var shapes = layout.Nodes.Select(n => n.Shape ).ToList();
            return shapes;

        }

        public IList<short> DrawGrid(
            IVisio.Master masterobj,
            VA.Drawing.Size cell_size,
            int cols,
            int rows)
        {
            
            //Create a new page to hold the grid
            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            
            using (var undoscope = application.CreateUndoScope())
            {
                var shapeids = VA.Layout.LayoutHelper.DrawGrid(page, masterobj, cell_size, cols, rows);
                return shapeids;
            }
        }

        public IVisio.Shape DrawNURBSCurve(IList<VA.Drawing.Point> controlpoints,
                                    IList<double> knots,
                                    IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {

                var page = application.ActivePage;
                var shape = page.DrawNURBS(controlpoints, knots, weights, degree);
                return shape;
            }
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = active_page.DrawRectangle(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = active_page.DrawLine(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(double x0, double y0, double x1, double y1)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = active_page.DrawOval(x0, y0, x1, y1);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(VA.Drawing.Point center, double radius)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var A = center.Add(-radius, -radius);
                var B = center.Add(radius, radius);
                var rect = new VA.Drawing.Rectangle(A, B);
                var active_page = application.ActivePage;
                var shape = active_page.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape DrawBezier(IEnumerable<VA.Drawing.Point> points)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = active_page.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape DrawPolyLine(IList<VA.Drawing.Point> points)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = active_page.DrawPolyline(points);
                return shape;
            }
        }

        public IVisio.Shape DrawPieSlice(VA.Drawing.Point center,
                                  double radius,
                                  double start_angle,
                                  double end_angle)
        {
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                var shape = DrawCommandsUtil.DrawPieSlice(active_page, center, radius, start_angle, end_angle);
                return shape;
            }
        }

        public IList<IVisio.Shape> DrawPieSlices(VA.Drawing.Point center,
                                          double radius,
                                          IList<double> values)
        {
            if (!this.Session.HasActiveDrawing())
            {
                return null;
            }

            var application = this.Session.VisioApplication;
            var page = application.ActivePage;
            var shapes = VA.Layout.LayoutHelper.DrawPieSlices(page, center, radius, values);
            return shapes;
        }
    }
}
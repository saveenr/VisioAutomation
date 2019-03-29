using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class DrawCommands : CommandSet
    {
        internal DrawCommands(Client client) :
            base(client)
        {

        }

        public VisioAutomation.SurfaceTarget GetActiveDrawingSurface()
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            var surf_application = cmdtarget.Application;
            var surf_window = surf_application.ActiveWindow;
            var surf_window_subtype = surf_window.SubType;

            // TODO: Revisit the logic here
            // TODO: And what about a selected shape as a surface?

            this._client.Output.WriteVerbose("Window SubType: {0}", surf_window_subtype);
            if (surf_window_subtype == 64)
            {
                this._client.Output.WriteVerbose("Window = Master Editing");
                var surf_master = (IVisio.Master)surf_window.Master;
                var surface = new VisioAutomation.SurfaceTarget(surf_master);
                return surface;

            }
            else
            {
                this._client.Output.WriteVerbose("Window = Page ");
                var surf_page = surf_application.ActivePage;
                var surface = new VisioAutomation.SurfaceTarget(surf_page);
                return surface;
            }
        }


        public IVisio.Shape DrawNurbsCurve(
            IList<VisioAutomation.Geometry.Point> controlpoints,
            IList<double> knots,
            IList<double> weights, int degree)
        {

            // flags:
            // None = 0,
            // IVisio.VisDrawSplineFlags.visSpline1D

            var cmdtarget = this._client.GetCommandTargetPage();

            var application = cmdtarget.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawNurbsCurve)))
            {

                var shape = cmdtarget.ActivePage.DrawNurbs(controlpoints, knots, weights, degree);
                return shape;
            }
        }

        public IVisio.Shape DrawRectangle(VisioAutomation.Geometry.Rectangle r)
        {
            var surface = this.GetActiveDrawingSurface();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawRectangle)))
            {
                var shape = surface.DrawRectangle(r.Left, r.Bottom, r.Right, r.Top);
                return shape;
            }
        }

        public IVisio.Shape DrawRectangle(double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.DrawRectangle(rect);
        }

        public IVisio.Shape DrawLine(double x0, double y0, double x1, double y1)
        {
            var p0 = new VisioAutomation.Geometry.Point(x0, y0);
            var p1 = new VisioAutomation.Geometry.Point(x1, y1);
            return this.DrawLine(p0, p1);
        }

        public IVisio.Shape DrawLine(VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1)
        {
            var surface = this.GetActiveDrawingSurface();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawLine)))
            {
                var shape = surface.DrawLine(p0,p1);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(VisioAutomation.Geometry.Rectangle rect)
        {
            var surface = this.GetActiveDrawingSurface();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawOval)))
            {
                var shape = surface.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.DrawOval(rect);
        }

        public IVisio.Shape DrawBezier(IEnumerable<VisioAutomation.Geometry.Point> points)
        {
            var surface = this.GetActiveDrawingSurface();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawOval)))
            {
                var shape = surface.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape DrawPolyLine(IList<VisioAutomation.Geometry.Point> points)
        {
            var surface = this.GetActiveDrawingSurface();
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawPolyLine)))
            {
                var shape = surface.DrawPolyLine(points);
                return shape;
            }
        }


        public void DuplicateShapes(int n)
        {
            var cmdtarget = this._client.GetCommandTargetDocument();

            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(n));
            }

            var window = cmdtarget.Application.ActiveWindow;
            var selection = window.Selection;
            if (selection.Count<1)
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = cmdtarget.Application;
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DuplicateShapes)))
            {
                var active_page = application.ActivePage;
                var new_shapes = DrawCommands._CreateDuplicates(active_page, selection[1], n);
            }
        }

        private static List<IVisio.Shape> _CreateDuplicates(IVisio.Page page,
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

        public List<IVisio.Shape> GetAllShapesOnActiveDrawingSurface()
        {
            var surface = this._client.ShapeSheet.GetShapeSheetSurface();
            var shapes = surface.Shapes;
            var list = new List<IVisio.Shape>();
            list.AddRange(shapes.ToEnumerable());
            return list;
        }
    }
}
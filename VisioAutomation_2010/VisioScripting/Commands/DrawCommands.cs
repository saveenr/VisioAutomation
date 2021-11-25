
using VisioAutomation.Extensions;


namespace VisioScripting.Commands
{
    public class DrawCommands : CommandSet
    {
        internal DrawCommands(Client client) :
            base(client)
        {

        }
        public IVisio.Shape DrawRectangle(VisioScripting.TargetPage targetpage, double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.DrawRectangle(targetpage,rect);
        }

        public IVisio.Shape DrawRectangle(VisioScripting.TargetPage targetpage, VisioAutomation.Geometry.Rectangle r)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawRectangle)))
            {
                var shape = targetpage.Page.DrawRectangle(r.Left, r.Bottom, r.Right, r.Top);
                return shape;
            }
        }


        public IVisio.Shape DrawLine(VisioScripting.TargetPage targetpage, double x0, double y0, double x1, double y1)
        {
            var p0 = new VisioAutomation.Geometry.Point(x0, y0);
            var p1 = new VisioAutomation.Geometry.Point(x1, y1);
            return this.DrawLine(targetpage, p0, p1);
        }

        public IVisio.Shape DrawLine(VisioScripting.TargetPage targetpage, VisioAutomation.Geometry.Point p0, VisioAutomation.Geometry.Point p1)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawLine)))
            {
                var shape = targetpage.Page.DrawLine(p0,p1);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(VisioScripting.TargetPage targetpage, VisioAutomation.Geometry.Rectangle rect)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            var surface = new VisioAutomation.SurfaceTarget(targetpage.Page);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawOval)))
            {
                var shape = targetpage.Page.DrawOval(rect);
                return shape;
            }
        }

        public IVisio.Shape DrawOval(VisioScripting.TargetPage targetpage, double x0, double y0, double x1, double y1)
        {
            var rect = new VisioAutomation.Geometry.Rectangle(x0, y0, x1, y1);
            return this.DrawOval(targetpage, rect);
        }

        public IVisio.Shape DrawBezier(VisioScripting.TargetPage targetpage, IEnumerable<VisioAutomation.Geometry.Point> points)
        {
            targetpage = targetpage.ResolveToPage(this._client);

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawOval)))
            {
                var shape = targetpage.Page.DrawBezier(points.ToList());
                return shape;
            }
        }

        public IVisio.Shape DrawPolyLine(VisioScripting.TargetPage targetpage, IList<VisioAutomation.Geometry.Point> points)
        {
            targetpage = targetpage.ResolveToPage(this._client);
            using (var undoscope = this._client.Undo.NewUndoScope(nameof(DrawPolyLine)))
            {
                var shape = targetpage.Page.DrawPolyLine(points);
                return shape;
            }
        }


        public void Duplicate(VisioScripting.TargetSelection targetselection,int n)
        {
            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException(nameof(n));
            }

            targetselection = targetselection.ResolveToSelection(this._client);

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            using (var undoscope = this._client.Undo.NewUndoScope(nameof(Duplicate)))
            {
                var app = targetselection.Selection.Application;
                var active_page = app.ActivePage;
                var new_shapes = DrawCommands._create_duplicates(active_page, targetselection.Selection[1], n);
            }
        }

        private static List<IVisio.Shape> _create_duplicates(IVisio.Page page,
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
    }
}
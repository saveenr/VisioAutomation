using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioAutomation.Scripting.Commands
{
    public class FormatCommands : CommandSet
    {
        public FormatCommands(Session session) :
            base(session)
        {

        }

        public void Set(IList<IVisio.Shape> target_shapes, VA.Format.ShapeFormatCells format)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count<1)
            {
                return;
            }

            var update = new VA.ShapeSheet.Update();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (int shapeid in shapeids)
            {
                update.SetFormulas((short)shapeid, format);
            }

            update.Execute(this.Session.VisioApplication.ActivePage);            
        }

        public IList<VA.Format.ShapeFormatCells> Get(IList<IVisio.Shape> target_shapes)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return new List<VA.Format.ShapeFormatCells>(0);
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var fmts = VA.Format.ShapeFormatCells.GetCells(this.Session.VisioApplication.ActivePage, shapeids);
            return fmts;
        }

        public void Duplicate(int n)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();
            
            if (n < 1)
            {
                throw new System.ArgumentOutOfRangeException("n");
            }
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            // TODO: Add ability to duplicate all the selected shapes, not just the first one
            // this dupicates exactly 1 shape N - times what it
            // it should do is duplicate all M selected shapes N times so that M*N shapes are created

            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication, string.Format("Duplicate Shape {0} Times",n)))
            {
                var active_window = application.ActiveWindow;
                var selection = active_window.Selection;
                var active_page = application.ActivePage;
                FormatCommands.CreateDuplicates(active_page, selection[1], n);
            }
        }

        private double? cached_size_width;
        private double? cached_size_height;

        /// <summary>
        /// Caches the resize (the results, not formulas) of a the first currently selected shape
        /// </summary>
        public void CopySize()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();
            
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            var shape = selection[1];

            var query = new VA.ShapeSheet.Query.CellQuery();
            var width_col = query.Columns.Add(VA.ShapeSheet.SRCConstants.Width, "Width");
            var height_col = query.Columns.Add(VA.ShapeSheet.SRCConstants.Height, "Height");
            var queryresults = query.GetResults<double>(shape);

            cached_size_width = queryresults[width_col.Ordinal];
            cached_size_height = queryresults[height_col.Ordinal];
        }

        public void PasteSize(IList<IVisio.Shape> target_shapes, bool paste_width, bool paste_height)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return;
            }

            if ((!cached_size_width.HasValue) && (!cached_size_height.HasValue))
            {
                return;
            }

            var update = new VA.ShapeSheet.Update();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (var shapeid in shapeids)
            {
                if (paste_width)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Width, cached_size_width.Value);
                }

                if (paste_height)
                {
                    update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.Height, cached_size_height.Value);
                }
            }

            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;
            update.Execute(active_page);
        }

        private readonly VA.Format.FormatPaintCache cache = new VA.Format.FormatPaintCache();

        public void CopyFormat()
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var allflags = this.cache.GetAllFormatPaintFlags();
            this.CopyFormat(null, allflags);
        }

        public void CopyFormat(IVisio.Shape target_shape, VA.Format.FormatCategory category)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shape = GetTargetShape(target_shape);
            if (shape == null)
            {
                return;
            }

            this.cache.CopyFormat(shape, category);
        }

        public void ClearFormatCache()
        {
            this.cache.Clear();
        }

        public void PasteFormat(IList<IVisio.Shape> target_shapes, VA.Format.FormatCategory category, bool apply_formulas)
        {
            this.CheckVisioApplicationAvailable();
            this.CheckActiveDrawingAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }
 
            var shapeids = target_shapes.Select(s=>s.ID).ToList();
            var application = this.Session.VisioApplication;
            var active_page = application.ActivePage;

            this.cache.PasteFormat(active_page, shapeids, category, apply_formulas);
        }

        public static IList<IVisio.Shape> CreateDuplicates(IVisio.Page page,
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

    }
}
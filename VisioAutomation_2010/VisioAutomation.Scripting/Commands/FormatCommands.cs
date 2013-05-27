using System.Collections.Generic;
using System.Xml.Linq;
using VisioAutomation.Extensions;
using System.Linq;
using VisioAutomation.Format;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

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
                return new List<ShapeFormatCells>(0);
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
                DrawCommandsUtil.CreateDuplicates(active_page, selection[1], n);
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
            var width_col = query.AddColumn(VA.ShapeSheet.SRCConstants.Width);
            var height_col = query.AddColumn(VA.ShapeSheet.SRCConstants.Height);
            var queryresults = query.GetResults<double>(shape);

            cached_size_width = queryresults[0, width_col];
            cached_size_height = queryresults[0, height_col];
        }

        /// <summary>
        /// Applies the cached size to the currently selected shapes. If no shapes are selected then nothing happens.
        /// If no size was cached then nothing happens.
        /// </summary>
        /// <param name="flags">Controls if either or both the width and height values are applied during the paste</param>
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

        private VA.Format.FormatPaintCache cache = new VA.Format.FormatPaintCache();

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
    }
}
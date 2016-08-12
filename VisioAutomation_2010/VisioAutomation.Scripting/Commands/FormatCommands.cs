using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheetQuery;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class FormatCommands : CommandSet
    {
        internal FormatCommands(Client client) :
            base(client)
        {

        }

        public void Set(IList<IVisio.Shape> target_shapes, Shapes.FormatCells format)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count<1)
            {
                return;
            }

            var update = new ShapeSheet.Update();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (int shapeid in shapeids)
            {
                update.SetFormulas((short)shapeid, format);
            }

            var application = this._client.Application.Get();
            update.Execute(application.ActivePage);            
        }

        public IList<Shapes.FormatCells> Get(IList<IVisio.Shape> target_shapes)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return new List<Shapes.FormatCells>(0);
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this._client.Application.Get();
            var fmts = Shapes.FormatCells.GetCells(application.ActivePage, shapeids);
            return fmts;
        }

        private double? cached_size_width;
        private double? cached_size_height;

        /// <summary>
        /// Caches the resize (the results, not formulas) of a the first currently selected shape
        /// </summary>
        public void CopySize()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (!this._client.Selection.HasShapes())
            {
                return;
            }

            var application = this._client.Application.Get();
            var active_window = application.ActiveWindow;
            var selection = active_window.Selection;
            var shape = selection[1];

            var query = new ShapeSheetQuery.Query();
            var width_col = query.AddCell(ShapeSheet.SRCConstants.Width, "Width");
            var height_col = query.AddCell(ShapeSheet.SRCConstants.Height, "Height");

            var ss = new ShapeSheetSurface(shape);

            var queryresults = query.GetResults<double>(ss);

            this.cached_size_width = queryresults.Cells[width_col];
            this.cached_size_height = queryresults.Cells[height_col];
        }

        public void PasteSize(IList<IVisio.Shape> target_shapes, bool paste_width, bool paste_height)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return;
            }

            if ((!this.cached_size_width.HasValue) && (!this.cached_size_height.HasValue))
            {
                return;
            }

            var update = new ShapeSheet.Update();
            var shapeids = shapes.Select(s => s.ID).ToList();

            foreach (var shapeid in shapeids)
            {
                if (paste_width)
                {
                    update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Width, this.cached_size_width.Value);
                }

                if (paste_height)
                {
                    update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.Height, this.cached_size_height.Value);
                }
            }

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;
            update.Execute(active_page);
        }

        private readonly FormatPaintCache cache = new FormatPaintCache();

        public void Copy()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var allflags = this.cache.GetAllFormatPaintFlags();
            this.Copy(null, allflags);
        }

        public void Copy(IVisio.Shape target_shape, FormatCategory category)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shape = this.GetTargetShape(target_shape);
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

        public void Paste(IList<IVisio.Shape> target_shapes, FormatCategory category, bool apply_formulas)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }
 
            var shapeids = target_shapes.Select(s=>s.ID).ToList();
            var application = this._client.Application.Get();
            var active_page = application.ActivePage;

            this.cache.PasteFormat(active_page, shapeids, category, apply_formulas);
        }
    }
}
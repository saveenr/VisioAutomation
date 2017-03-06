using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Scripting.FormatPaint;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Commands
{
    public class FormatCommands : CommandSet
    {
        internal FormatCommands(Client client) :
            base(client)
        {

        }

        public void Set(TargetShapes targets, Shapes.ShapeFormatCells format)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            
            if (targets.Shapes.Count<1)
            {
                return;
            }

            var writer = new SidSrcWriter();
            var shapeids = targets.Shapes.Select(s => s.ID).ToList();

            foreach (int shapeid in shapeids)
            {
                format.SetFormulas((short)shapeid, writer);
            }

            var application = this._client.Application.Get();

            writer.Commit(application.ActivePage);            
        }

        public IList<Shapes.ShapeFormatCells> Get(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            
            if (targets.Shapes.Count < 1)
            {
                return new List<Shapes.ShapeFormatCells>(0);
            }

            var shapeids = targets.Shapes.Select(s => s.ID).ToList();
            var application = this._client.Application.Get();
            var fmts = Shapes.ShapeFormatCells.GetCells(application.ActivePage, shapeids);
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

            var query = new ShapeSheetQuery();
            var width_col = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, "Width");
            var height_col = query.AddCell(VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, "Height");

            var queryresults = query.GetResults<double>(shape);

            this.cached_size_width = queryresults.Cells[width_col];
            this.cached_size_height = queryresults.Cells[height_col];
        }

        public void PasteSize(TargetShapes targets, bool paste_width, bool paste_height)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            if ((!this.cached_size_width.HasValue) && (!this.cached_size_height.HasValue))
            {
                return;
            }

            var writer = new SidSrcWriter();
            var shapeids = targets.Shapes.Select(s => s.ID).ToList();

            foreach (var shapeid in shapeids)
            {
                if (paste_width)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormWidth, this.cached_size_width.Value);
                }

                if (paste_height)
                {
                    writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.XFormHeight, this.cached_size_height.Value);
                }
            }

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;

            writer.Commit(active_page);
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

            var targets = new VisioAutomation.Scripting.TargetShapes( target_shape );
            var shapes = targets.ResolveShapes(this._client);
            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var shape = shapes.Shapes[0];

            this.cache.CopyFormat(shape, category);
        }

        public void ClearFormatCache()
        {
            this.cache.Clear();
        }

        public void Paste(TargetShapes targets, FormatCategory category, bool apply_formulas)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }
 
            var shapeids = targets.Shapes.Select(s=>s.ID).ToList();
            var application = this._client.Application.Get();
            var active_page = application.ActivePage;

            this.cache.PasteFormat(active_page, shapeids, category, apply_formulas);
        }
    }
}
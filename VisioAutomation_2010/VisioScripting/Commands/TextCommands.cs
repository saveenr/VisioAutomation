using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class TextCommands : CommandSet
    {
        internal TextCommands(Client client) :
            base(client)
        {

        }

        public void Set(VisioScripting.Models.TargetShapes targets, IList<string> texts)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            if (texts == null || texts.Count < 1)
            {
                return;
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set Shape Text"))
            {
                // Apply text to each shape
                // if there are fewer texts than shapes then
                // start reusing the texts from the beginning

                int count = 0;
                foreach (var shape in targets.Shapes)
                {
                    string text = texts[count%texts.Count];
                    if (text != null)
                    {
                        shape.Text = text;
                    }
                    count++;
                }
            }
        }

        public List<string> Get(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = targets.Shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void ToogleCase(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);
            
            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Toggle Shape Text Case"))
            {
                var shapeids = targets.Shapes.Select(s => s.ID).ToList();

                var page = application.ActivePage;
                // Store all the formatting
                var formats = VisioAutomation.Text.TextFormat.GetFormat(page, shapeids);

                // Change the text - this will wipe out all the character and paragraph formatting
                foreach (var shape in targets.Shapes)
                {
                    string t = shape.Text;
                    if (t.Length < 1)
                    {
                        continue;
                    }
                    shape.Text = VisioScripting.Helpers.TextHelper.toggle_case(t);
                }

                // Now restore all the formatting - based on any initial formatting from the text

                var writer = new SidSrcWriter();
                for (int i = 0; i < targets.Shapes.Count; i++)
                {
                    var format = formats[i];

                    if (format.CharacterFormats.Count > 0)
                    {
                        var fmt = format.CharacterFormats[0];
                        fmt.SetFormulas((short)shapeids[i], writer, 0);
                    }

                    if (format.ParagraphFormats.Count > 0)
                    {
                        var fmt = format.ParagraphFormats[0];
                        fmt.SetFormulas((short)shapeids[i], writer, 0);
                    }
                }

                writer.Commit(page);
            }
        }

        public void SetFont(VisioScripting.Models.TargetShapes targets, string fontname)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var active_document = application.ActiveDocument;
            var active_doc_fonts = active_document.Fonts;
            var font = active_doc_fonts[fontname];
            var page = this._client.Page.Get();

            var cells = new VisioAutomation.Text.CharacterCells();
            cells.Font = font.ID;

            this._client.ShapeSheet.__SetCells(targets, cells, page);
        }



        public List<VisioAutomation.Text.TextFormat> GetFormat(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<VisioAutomation.Text.TextFormat>(0);
            }

            var selection = this._client.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this._client.Application.Get();
            var formats = VisioAutomation.Text.TextFormat.GetFormat(application.ActivePage, shapeids);
            return formats;
        }

        public void MoveTextToBottom(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var writer = new SidSrcWriter();
            foreach (var shape in targets.Shapes)
            {
                if (0 ==
                    shape.RowExists[
                        (short)IVisio.VisSectionIndices.visSectionObject, (short)IVisio.VisRowIndices.visRowTextXForm,
                        (short)IVisio.VisExistsFlags.visExistsAnywhere])
                {
                    shape.AddRow((short)IVisio.VisSectionIndices.visSectionObject, (short)IVisio.VisRowIndices.visRowTextXForm, (short)IVisio.VisRowTags.visTagDefault);

                }
            }

            var application = this._client.Application.Get();
            var shapeids = targets.Shapes.Select(s => s.ID);
            foreach (int shapeid in shapeids)
            {
                writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.TextXFormHeight, "Height*0");
                writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.TextXFormPinY, "Height*0");
                writer.SetFormula((short)shapeid, VisioAutomation.ShapeSheet.SrcConstants.TextBlockVerticalAlign, "0");
            }

            writer.Commit(application.ActivePage);
        }

        public void SetTextWrapping(VisioScripting.Models.TargetShapes targets, bool wrap)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2D(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();
            var active_page = this._client.Page.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Set Text Wrapping"))
            {
                VisioScripting.Helpers.TextHelper.set_text_wrapping(active_page, shapeids, wrap);
            }
        }

        public void FitShapeToText(VisioScripting.Models.TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            var shapes = targets.ResolveShapes2D(this._client);

            if (shapes.Shapes.Count < 1)
            {
                return;
            }

            var application = this._client.Application.Get();
            var active_page = application.ActivePage;
            var shapeids = shapes.Shapes.Select(s => s.ID).ToList();

            using (var undoscope = this._client.Application.NewUndoScope("Fit Shape To Text"))
            {
                // Calculate the new sizes for each shape
                var new_sizes = new List<VisioAutomation.Drawing.Size>(shapeids.Count);
                foreach (var shape in shapes.Shapes)
                {
                    var text_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightText).Size;
                    var wh_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH).Size;

                    double max_w = System.Math.Max(text_bounding_box.Width, wh_bounding_box.Width);
                    double max_h = System.Math.Max(text_bounding_box.Height, wh_bounding_box.Height);
                    var max_size = new VisioAutomation.Drawing.Size(max_w, max_h);
                    new_sizes.Add(max_size);
                }

                var src_width = VisioAutomation.ShapeSheet.SrcConstants.XFormWidth;
                var src_height = VisioAutomation.ShapeSheet.SrcConstants.XFormHeight;

                var writer = new SidSrcWriter();
                for (int i = 0; i < new_sizes.Count; i++)
                {
                    var shapeid = shapeids[i];
                    var new_size = new_sizes[i];
                    writer.SetFormula((short)shapeid, src_width, new_size.Width);
                    writer.SetFormula((short)shapeid, src_height, new_size.Height);
                }

                writer.Commit(active_page);
            }
        }
    }
}



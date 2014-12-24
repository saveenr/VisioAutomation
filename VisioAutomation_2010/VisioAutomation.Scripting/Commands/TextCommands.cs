using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class TextCommands: CommandSet
    {
        public TextCommands(Client client) :
            base(client)
        {

        }

        public void Set(IList<IVisio.Shape> target_shapes, IList<string> texts)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            if (texts == null || texts.Count<1)
            {
                // do nothing
                return;
            }
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count<1)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Set Shape Text"))
            {
                int numtexts = texts.Count;
                for (int i=0;i<shapes.Count;i++)
                {
                    var shape = shapes[i];
                    var text = texts[i % numtexts];
                    shape.Text = text;
                }
            }
        }

        public IList<string> Get(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void ToogleCase(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication,"Toggle Shape Text Case"))
            {
                var shapeids = shapes.Select(s => s.ID).ToList();

                var page = this.Client.VisioApplication.ActivePage;
                // Store all the formatting
                var formats = VA.Text.TextFormat.GetFormat(page, shapeids);

                // Change the text - this will wipe out all the character and paragraph formatting
                foreach (var shape in shapes)
                {
                    string t = shape.Text;
                    if (t.Length < 1)
                    {
                        continue;
                    }
                    shape.Text = TextCommandsUtil.toggle_case(t);
                }

                // Now restore all the formatting - based on any initial formatting from the text

                var update = new VA.ShapeSheet.Update();
                for (int i = 0; i < shapes.Count; i++)
                {
                    var format = formats[i];

                    if (format.CharacterFormats.Count>0)
                    {
                        var fmt = format.CharacterFormats[0];
                        update.SetFormulas((short) shapeids[i], fmt, 0);
                    }

                    if (format.ParagraphFormats.Count > 0)
                    {
                        var fmt = format.ParagraphFormats[0];
                        update.SetFormulas((short)shapeids[i], fmt, 0);
                    }
                }

                update.Execute(page);
            }
        }

        public void SetFont(IList<IVisio.Shape> target_shapes, string fontname)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }
            var application = this.Client.VisioApplication;
            var active_document = application.ActiveDocument;
            var active_doc_fonts = active_document.Fonts;
            var font = active_doc_fonts[fontname];
            IVisio.VisGetSetArgs flags=0;
            var srcs = new[] {VA.ShapeSheet.SRCConstants.CharFont};
            var formulas = new[] { font.ID.ToString() };
            this.Client.ShapeSheet.SetFormula(target_shapes, srcs, formulas, flags);
        }

        public IList<VA.Text.TextFormat> GetFormat(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<VA.Text.TextFormat>(0);
            }

            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Client.VisioApplication;
            var formats = VA.Text.TextFormat.GetFormat(application.ActivePage, shapeids);
            return formats;
        }

        public void MoveTextToBottom(IList<IVisio.Shape> target_shapes)
        {
            this.AssertApplicationAvailable();
            this.AssertDocumentAvailable();

            var shapes = GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return ;
            }

            var update = new VA.ShapeSheet.Update();
            foreach (var shape in shapes)
            {
                if (0 ==
                    shape.RowExists[
                        (short) IVisio.VisSectionIndices.visSectionObject, (short) IVisio.VisRowIndices.visRowTextXForm,
                        (short) IVisio.VisExistsFlags.visExistsAnywhere])
                {
                    shape.AddRow((short)IVisio.VisSectionIndices.visSectionObject, (short)IVisio.VisRowIndices.visRowTextXForm, (short)IVisio.VisRowTags.visTagDefault); 
                    
                }
            }

            var application = this.Client.VisioApplication;
            var shapeids = shapes.Select(s=>s.ID);
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.TxtHeight, "Height*0"); 
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.TxtPinY, "Height*0"); 
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.VerticalAlign, "0");
            } 
            var active_page = application.ActivePage; 
            update.Execute(active_page);
        }

        public void SetTextWrapping(bool wrap)
        {
            if (!this.Client.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(application,"SetTextWrapping"))
            {
                var active_page = application.ActivePage;
                TextCommandsUtil.set_text_wrapping(active_page, shapeids, wrap);
            }
        }

        public void FitShapeToText()
        {


            if (!this.Client.HasSelectedShapes())
            {
                return;
            }
            var shapes_2d = this.Client.Selection.GetShapes();
            var application = this.Client.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(application,"FitShapeToText"))
            {
                var active_page = application.ActivePage;
                VA.Text.TextHelper.FitShapeToText(active_page, shapes_2d);
            }
        }
    }
}


namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        public static void FitShapeToText(IVisio.Page page, IEnumerable<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            // Calculate the new sizes for each shape
            var new_sizes = new List<VA.Drawing.Size>(shapeids.Count);
            foreach (var shape in shapes)
            {
                var text_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightText).Size;
                var wh_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH).Size;

                double max_w = System.Math.Max(text_bounding_box.Width, wh_bounding_box.Width);
                double max_h = System.Math.Max(text_bounding_box.Height, wh_bounding_box.Height);
                var max_size = new VA.Drawing.Size(max_w, max_h);
                new_sizes.Add(max_size);
            }

            var src_width = VA.ShapeSheet.SRCConstants.Width;
            var src_height = VA.ShapeSheet.SRCConstants.Height;

            var update = new VA.ShapeSheet.Update();
            for (int i = 0; i < new_sizes.Count; i++)
            {
                var shapeid = shapeids[i];
                var new_size = new_sizes[i];
                update.SetFormula((short) shapeid, src_width, new_size.Width);
                update.SetFormula((short) shapeid, src_height, new_size.Height);
            }

            update.Execute(page);
        }


        public static IVisio.Font TryGetFont(IVisio.Fonts fonts, string name)
        {
            try
            {
                var font = fonts[name];
                return font;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }
    }
}



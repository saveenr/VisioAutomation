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
           Set(target_shapes, texts, false);
       }

        public void Set(IList<IVisio.Shape> target_shapes, IList<string> texts, bool set_Name)
       {
           this.Client.Application.AssertApplicationAvailable();
           this.Client.Document.AssertDocumentAvailable();

           if (texts == null || texts.Count < 1)
           {
               // do nothing
               return;
           }

           var shapes = this.GetTargetShapes(target_shapes);
           if (shapes.Count < 1)
           {
               return;
           }

           using (var undoscope = new VA.Application.UndoScope(this.Client.VisioApplication, "Set Shape Text"))
           {
               int numtexts = texts.Count;
               for (int i = 0; i < shapes.Count; i++)
               {
                   var shape = shapes[i];
                   var text = texts[i % numtexts];
                   shape.Text = text;
                   if (set_Name)
                   {
                       shape.Name = text;
                   }
               }
           }
       }

        public IList<string> Get(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void ToogleCase(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();
            
            var shapes = this.GetTargetShapes(target_shapes);

            if (shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = new Application.UndoScope(this.Client.VisioApplication,"Toggle Shape Text Case"))
            {
                var shapeids = shapes.Select(s => s.ID).ToList();

                var page = this.Client.VisioApplication.ActivePage;
                // Store all the formatting
                var formats = Text.TextFormat.GetFormat(page, shapeids);

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

                var update = new ShapeSheet.Update();
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
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

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
            var srcs = new[] {ShapeSheet.SRCConstants.CharFont};
            var formulas = new[] { font.ID.ToString() };
            this.Client.ShapeSheet.SetFormula(target_shapes, srcs, formulas, flags);
        }

        public IList<Text.TextFormat> GetFormat(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<Text.TextFormat>(0);
            }

            var selection = this.Client.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Client.VisioApplication;
            var formats = Text.TextFormat.GetFormat(application.ActivePage, shapeids);
            return formats;
        }

        public void MoveTextToBottom(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes(target_shapes);
            if (shapes.Count < 1)
            {
                return ;
            }

            var update = new ShapeSheet.Update();
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
                update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.TxtHeight, "Height*0"); 
                update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.TxtPinY, "Height*0"); 
                update.SetFormula((short)shapeid, ShapeSheet.SRCConstants.VerticalAlign, "0");
            } 
            var active_page = application.ActivePage; 
            update.Execute(active_page);
        }

        public void SetTextWrapping(IList<IVisio.Shape> target_shapes,bool wrap)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes2D(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var shapeids = shapes.Select(s => s.ID).ToList();
            var application = this.Client.VisioApplication;
            using (var undoscope = new Application.UndoScope(application,"SetTextWrapping"))
            {
                var active_page = application.ActivePage;
                TextCommandsUtil.set_text_wrapping(active_page, shapeids, wrap);
            }
        }

        public void FitShapeToText(IList<IVisio.Shape> target_shapes)
        {
            this.Client.Application.AssertApplicationAvailable();
            this.Client.Document.AssertDocumentAvailable();

            var shapes = this.GetTargetShapes2D(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            var application = this.Client.VisioApplication;
            var active_page = application.ActivePage;
            var shapeids = shapes.Select(s => s.ID).ToList();

            using (var undoscope = new Application.UndoScope(application, "FitShapeToText"))
            {
                // Calculate the new sizes for each shape
                var new_sizes = new List<Drawing.Size>(shapeids.Count);
                foreach (var shape in shapes)
                {
                    var text_bounding_box = shape.GetBoundingBox(Microsoft.Office.Interop.Visio.VisBoundingBoxArgs.visBBoxUprightText).Size;
                    var wh_bounding_box = shape.GetBoundingBox(Microsoft.Office.Interop.Visio.VisBoundingBoxArgs.visBBoxUprightWH).Size;

                    double max_w = System.Math.Max(text_bounding_box.Width, wh_bounding_box.Width);
                    double max_h = System.Math.Max(text_bounding_box.Height, wh_bounding_box.Height);
                    var max_size = new Drawing.Size(max_w, max_h);
                    new_sizes.Add(max_size);
                }

                var src_width = ShapeSheet.SRCConstants.Width;
                var src_height = ShapeSheet.SRCConstants.Height;

                var update = new ShapeSheet.Update();
                for (int i = 0; i < new_sizes.Count; i++)
                {
                    var shapeid = shapeids[i];
                    var new_size = new_sizes[i];
                    update.SetFormula((short)shapeid, src_width, new_size.Width);
                    update.SetFormula((short)shapeid, src_height, new_size.Height);
                }

                update.Execute(active_page);
            }
        }
    }
}



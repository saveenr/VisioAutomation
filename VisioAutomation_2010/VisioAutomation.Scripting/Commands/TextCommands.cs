using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class TextCommands: CommandSet
    {
        public TextCommands(Session session) :
            base(session)
        {

        }

        public void SetText(IList<IVisio.Shape> target_shapes, string text)
        {
            this.SetText(target_shapes, new string [] { text });
        }

        public void SetText(IList<IVisio.Shape> target_shapes, IEnumerable<string> texts)
        {
            var shapes = this.get_target_shapes(target_shapes);
            if (shapes.Count<1)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set Shape Tex"))
            {
                var values = texts.ToList();

                for (int i=0;i<shapes.Count;i++)
                {
                    var shape = shapes[i];
                    var text = values[i%values.Count];
                    shape.Text = text;
                }
            }
        }

        public IList<string> GetText(IList<IVisio.Shape> target_shapes)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<string>(0);
            }

            var texts = shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void ToogleCase(IList<IVisio.Shape> target_shapes)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Toggle Shape Text Case"))
            {
                var shapeids = shapes.Select(s => s.ID).ToList();

                var page = this.Session.VisioApplication.ActivePage;
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
                        update.SetFormulasForRow((short) shapeids[i], fmt, (short)0);
                    }

                    if (format.ParagraphFormats.Count > 0)
                    {
                        var fmt = format.ParagraphFormats[0];
                        update.SetFormulasForRow((short)shapeids[i], fmt, (short)0);
                    }
                }

                update.Execute(page);
            }
        }

        //TODO: Make this support an input list
        public void SetTextWrapping(bool wrap)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Set Text Wrapping"))
            {
                var active_page = application.ActivePage;
                TextCommandsUtil.set_text_wrapping(active_page, shapeids, wrap);
            }
        }

        //TODO: Make this support an input list
        public void FitShapeToText()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes_2d = this.Session.Selection.EnumShapes2D().ToList();
            var application = this.Session.VisioApplication;
            using (var undoscope = new VA.Application.UndoScope(this.Session.VisioApplication,"Fit Shape To Text"))
            {
                var active_page = application.ActivePage;
                VA.Text.TextHelper.FitShapeToText(active_page, shapes_2d);
            }
        }

        //TODO: Make this support an input list
        public void MoveTextToBottom()
        {
            // http://www.visguy.com/2007/11/07/text-to-the-bottom-of-the-shape/

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var application = this.Session.VisioApplication;
            var active_window = application.ActiveWindow;
            var sel = active_window.Selection;
            var shapes = this.Session.Selection.EnumShapes().ToList();
            var update = new VA.ShapeSheet.Update();

            foreach (var shape in shapes)
            {
                if (0 == shape.RowExists[(short)IVisio.VisSectionIndices.visSectionObject, (short)IVisio.VisRowIndices.visRowTextXForm, (short)IVisio.VisExistsFlags.visExistsAnywhere])
                {
                    shape.AddRow((short)IVisio.VisSectionIndices.visSectionObject,
                                 (short)IVisio.VisRowIndices.visRowTextXForm,
                                 (short)IVisio.VisRowTags.visTagDefault);
                }
            }

            var shapeids = sel.GetIDs();
            foreach (int shapeid in shapeids)
            {
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.TxtHeight, "Height*0");
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.TxtPinY, "Height*0");
                update.SetFormula((short)shapeid, VA.ShapeSheet.SRCConstants.VerticalAlign, "0");
            }

            var active_page = application.ActivePage;
            update.Execute(active_page);
        }


        private static IVisio.Font TryGetFont(IVisio.Fonts fonts, string name)
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

        public void SetStyleProperties(string stylename, string fontname)
        {
            if (!this.Session.HasActiveDrawing)
            {
                return;
            }

            var doc = this.Session.VisioApplication.ActiveDocument;
            var styles = doc.Styles;
            var style = styles.ItemU[stylename];

            if (fontname != null)
            {
                var font = TryGetFont(doc.Fonts, fontname);

                if (font == null)
                {
                    var msg = "No such font: " + fontname;
                    throw new System.ArgumentException(msg, "fontname");
                }
                var src_Char_Font = VA.ShapeSheet.SRCConstants.Char_Font;

                var cell_font = style.CellsSRC[src_Char_Font.Section, src_Char_Font.Row, src_Char_Font.Cell];
                cell_font.FormulaU = font.ID.ToString(System.Globalization.CultureInfo.InvariantCulture);                
            }
        }

        //TODO: Make this support an input list
        public void SetFont(string fontname)
        {
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var active_doc_fonts = active_document.Fonts;
            var font = active_doc_fonts[fontname];
            var fontids = new[] {font.ID.ToString()};
            IVisio.VisGetSetArgs flags=0;
            this.Session.ShapeSheet.SetFormula(new[] { VA.ShapeSheet.SRCConstants.Char_Font }, fontids, flags);
        }

        public IList<VA.Text.TextFormat> GetFormat(IList<IVisio.Shape> target_shapes)
        {
            var shapes = get_target_shapes(target_shapes);
            if (shapes.Count < 1)
            {
                return new List<VA.Text.TextFormat>(0);
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            var formats = VA.Text.TextFormat.GetFormat(application.ActivePage, shapeids);
            return formats;
        }
    }
}
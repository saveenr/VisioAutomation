using System;
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

        public void SetText(string text)
        {
            var texts = new string[] {text};
            SetText(texts);
        }

        public void SetText(IEnumerable<string> texts)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();

            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
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

        public IList<string> GetText()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new List<string>(0);
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var texts = shapes.Select(s => s.Text).ToList();
            return texts;
        }

        public void ToogleCase()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            int rounding = 0;
            var shapes = this.Session.Selection.EnumShapes().ToList();
            var application = this.Session.VisioApplication;
            var src_charstyle = VA.ShapeSheet.SRCConstants.Char_Style;

            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    var s = shape; // to prevent Access to Modified Closure warning
                    var tf = VA.Text.TextFormat.GetFormat(s);
                    var textruns = tf.CharacterTextRuns;
                    var nocast = (short)IVisio.VisUnitCodes.visNoCast;
                    var textstyles = textruns
                        .Select(
                            tr =>
                                {
                                    var c = s.GetCell(src_charstyle);
                                    return (short) c.ResultInt[nocast, (short) rounding];
                                }
                        ).ToList();

                    string t = s.Text;
                    if (t.Length < 1)
                    {
                        continue;
                    }
                    s.Text = TextCommandsUtil.toggle_case(t);

                    foreach (var tr in textruns)
                    {
                        var chars = s.Characters;
                        chars.Begin = tr.Begin;
                        chars.End = tr.End;
                        var cellindex = src_charstyle.Cell;
                        chars.CharProps[cellindex] =  textstyles[tr.Index];
                    }
                }
            }
        }

        public void InsertField(VA.Text.Markup.Field field, int start, int end)
        {
            if (start < 0)
            {
                throw new ArgumentOutOfRangeException("end", "must be greater than or equal to zero");
            }

            if (end < start)
            {
                throw new ArgumentOutOfRangeException("end", "must be greater than or equal to start");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    var c = shape.Characters;
                    c.Begin = start;
                    c.End = end;
                    c.AddField((short)field.Category, (short)field.Code, (short)field.Format);
                }
            }
        }

        public void InsertCustomField(int start, int end, string formula, IVisio.VisFieldFormats format)
        {
            if (start < 0)
            {
                throw new ArgumentOutOfRangeException("end", "must be greater than or equal to zero");
            }

            if (end < start)
            {
                throw new ArgumentOutOfRangeException("end", "must be greater than or equal to start");
            }

            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes = this.Session.Selection.EnumShapes().ToList();
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                foreach (var shape in shapes)
                {
                    var c = shape.Characters;
                    c.Begin = start;
                    c.End = end;
                    c.AddCustomFieldU(formula,(short)format);
                }
            }
        }

        public void SetTextWrapping(bool wrap)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                TextCommandsUtil.set_text_wrapping(active_page, shapeids, wrap);
            }
        }

        public void FitShapeToText()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var shapes_2d = this.Session.Selection.EnumShapes2D().ToList();
            var application = this.Session.VisioApplication;
            using (var undoscope = application.CreateUndoScope())
            {
                var active_page = application.ActivePage;
                VA.Text.TextHelper.FitShapeToText(active_page, shapes_2d);
            }
        }

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
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

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

        public void StripWhiteSpace()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }

            var all_shapes = this.Session.Selection.EnumShapes().ToList();

            foreach (var shape in all_shapes)
            {
                var original_text = shape.Text;
                var stripped_text = original_text.Trim();
                if (original_text.Length != stripped_text.Length)
                {
                    shape.Text = stripped_text;
                }
            }
        }

        public void IncreaseTextSize()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }
            this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdSetCharSizeUp);
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
                var font = VA.Text.TextHelper.TryGetFont(doc.Fonts, fontname);

                if (font == null)
                {
                    var msg = "No such font: " + fontname;
                    throw new ArgumentException(msg, "fontname");
                }
                var src_Char_Font = VA.ShapeSheet.SRCConstants.Char_Font;

                var cell_font = style.CellsSRC[src_Char_Font.Section, src_Char_Font.Row, src_Char_Font.Cell];
                cell_font.FormulaU = font.ID.ToString(System.Globalization.CultureInfo.InvariantCulture);                
            }
        }

        public void SetTextFont(string fontname)
        {
            var fontnames = new string[] {fontname};
            var application = this.Session.VisioApplication;
            var active_document = application.ActiveDocument;
            var active_doc_fonts = active_document.Fonts;
            var fonts = fontnames.Select(v => active_doc_fonts[v]);
            var fontids = fonts.Select(f => f.ID.ToString()).ToList();
            IVisio.VisGetSetArgs flags=0;
            this.Session.ShapeSheet.SetFormula(new[] { VA.ShapeSheet.SRCConstants.Char_Font }, fontids, flags);
        }

        public void DecreaseTextSize()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return;
            }
            this.Session.VisioApplication.DoCmd((short)IVisio.VisUICmds.visCmdSetCharSizeDown);
        }

        public IList<VA.Text.TextFormat> GetTextFormat()
        {
            if (!this.Session.HasSelectedShapes())
            {
                return new List<VA.Text.TextFormat>(0);
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            var formats = VA.Text.TextFormat.GetFormat(application.ActivePage, shapeids);
            return formats;
        }

        public void SetTextFormat(VA.Text.CharacterFormatCells charfmt, VA.Text.ParagraphFormatCells parafmt)
        {
            if (!this.Session.HasSelectedShapes())
            {
                return ;
            }

            var selection = this.Session.Selection.Get();
            var shapeids = selection.GetIDs();
            var application = this.Session.VisioApplication;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            if (charfmt != null)
            {
                foreach (int shapeid in shapeids)
                {
                    charfmt.Apply(update,(short)shapeid,0);
                }
            }

            if (parafmt != null)
            {
                foreach (int shapeid in shapeids)
                {
                    parafmt.Apply(update, (short)shapeid, 0);
                }
            }

            update.Execute(application.ActivePage);
        }
    }
}
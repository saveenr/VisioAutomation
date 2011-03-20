using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        private static readonly VA.Text.CharacterFormatQuery charquery = new VA.Text.CharacterFormatQuery();

        public static void SetTextFormatFields(IVisio.Shape shape, string fmt, params object[] fields)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (fields == null)
            {
                throw new ArgumentNullException("fields");
            }

            var t_string = typeof (string);
            var t_field_element = typeof (VA.Text.Markup.Field);

            for (int i = 0; i < fields.Length; i++)
            {
                object field = fields[i];
                if (field == null)
                {
                    string msg = String.Format("Field value {0} is null", i);
                    throw new ArgumentException(msg);
                }

                var ft = field.GetType();

                if ((ft != t_string) && (ft != t_field_element))
                {
                    string msg = String.Format("Field value {0} is must be {1} or {2}. Instead it is {3}", i,
                                               t_string.Name, t_field_element.Name, ft.Name);
                    throw new ArgumentException(msg);
                }
            }

            var fmtparse = new VA.FormatStringParser(fmt);
            var unique_indices = fmtparse.Segments.Select(f => f.Index).Distinct().ToList();
            if (unique_indices.Count > fields.Length)
            {
                throw new ArgumentOutOfRangeException("fmt", "index out of range for number of insertions");
            }

            // Set the text
            shape.Text = fmt;

            // then Insert the fields from back to front
            for (int i = (fmtparse.Segments.Count - 1); i >= 0; i--)
            {
                var fmt_seg = fmtparse.Segments[i];
                var field_index = fmt_seg.Index;
                object field = fields[field_index];

                var chars = shape.Characters;
                chars.Begin = fmt_seg.Start;
                chars.End = fmt_seg.End;

                var ft = field.GetType();
                if (t_string == ft)
                {
                    // it must be a formula
                    string formula = (string) field;
                    chars.AddCustomFieldU(formula, (short) IVisio.VisFieldFormats.visFmtNumGenNoUnits);
                }
                else if (t_field_element == ft)
                {
                    var field_f = (VA.Text.Markup.Field) field;
                    chars.AddField((short) field_f.Category, (short) field_f.Code, (short) field_f.Format);
                }
                else
                {
                    string msg = String.Format("Unsupported field type {0} for field {1}", ft.Name, i);
                    throw new AutomationException(msg);
                }
            }
        }


        /// <summary>
        /// Tests whether a font is available to the Visio application. The method is not case sensitive
        /// </summary>
        /// <param name="fonts">Visio Fonts Object</param>
        /// <param name="fontname">the name of the font to find.</param>
        /// <returns>null if the font cannot be found, otherwise the font object</returns>
        public static IVisio.Font FindFontWithName(IVisio.Fonts fonts, string fontname)
        {
            if (fontname == null)
            {
                throw new ArgumentNullException("fontname");
            }

            if (String.IsNullOrEmpty(fontname))
            {
                throw new ArgumentException("fontname");
            }

            foreach (var f in fonts.AsEnumerable())
            {
                if (String.Compare(f.Name, fontname, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    return f;
                }
            }

            return null;
        }

        internal static IVisio.Characters SetRangeParagraphProps(IVisio.Shape shape, short cell, int value, int begin,
                                                                 int end)
        {
            var chars = shape.Characters;
            chars.Begin = begin;
            chars.End = end;
            chars.ParaProps[cell] = (short) value;
            return chars;
        }

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
                var max_size = VA.Drawing.DrawingUtil.Max(text_bounding_box, wh_bounding_box);
                new_sizes.Add(max_size);
            }

            var src_width = VA.ShapeSheet.SRCConstants.Width;
            var src_height = VA.ShapeSheet.SRCConstants.Height;

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            for (int i = 0; i < new_sizes.Count; i++)
            {
                var shapeid = shapeids[i];
                var new_size = new_sizes[i];
                update.SetFormula((short) shapeid, src_width, new_size.Width);
                update.SetFormula((short) shapeid, src_height, new_size.Height);
            }

            update.Execute(page);
        }

        internal enum rangetype
        {
            Paragraph,
            Character
        }

        internal static void SetRangeProps<T>(IVisio.Shape shape, VA.ShapeSheet.CellData<T> f,
                                              IVisio.VisCellIndices cell, int value, int begin, int end,
                                              ref short rownum, ref IVisio.Characters chars, rangetype rt)
        {
            if (f.Formula.HasValue)
            {
                var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;
                chars = shape.Characters;
                chars.Begin = begin;
                chars.End = end;

                if (rt == rangetype.Character)
                {
                    chars.CharProps[(short) cell] = (short) value;
                    rownum = chars.CharPropsRow[(short) default_chars_bias];
                }
                else if (rt == rangetype.Paragraph)
                {
                    chars.ParaProps[(short) cell] = (short) value;
                    rownum = chars.ParaPropsRow[(short) default_chars_bias];
                }
                else
                {
                    throw new ArgumentOutOfRangeException("rangetype");
                }

                if (rownum < 0)
                {
                    throw new VA.AutomationException("Failed to set CharPropsRow");
                }
            }
        }


        public static IList<TextRun> GetTextRuns(
            IVisio.Shape shape,
            IVisio.VisRunTypes runtype,
            bool collect_text)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var runs = new List<TextRun>();

            // based on this example: http://blogs.msdn.com/visio/archive/2006/08/18/704811.aspx
            // Get the Characters object representing the shape text
            var chars = shape.Characters;
            int num_chars = chars.CharCount;
            int run_begin = 0;
            int run_end = 1;

            int index = 0;

            // Find the beginning point and end point of every text run in the shape
            for (int c = 0; c < num_chars; c = run_end)
            {
                // Set the begin and end of the Characters object to the current position
                chars.Begin = c;
                chars.End = c + 1;

                // Get the beginning and end of this character run
                run_begin = chars.RunBegin[(short) runtype];
                run_end = chars.RunEnd[(short) runtype];

                // Set the begin and end of the Characters object to this run
                chars.Begin = run_begin;
                chars.End = run_end;

                // Record the text in this run
                string t = null;
                if (collect_text)
                {
                    t = chars.TextAsString;
                }

                var textrun = new TextRun(index, run_begin, run_end, t);
                index++;
                runs.Add(textrun);

                // As the for loop proceeds, c is set to the end of the current run
            }

            return runs;
        }


        public static void SetFormat(IVisio.Shape shape, VA.Text.CharacterFormatCells fmt)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            int rowcount = shape.RowCount[(short) IVisio.VisSectionIndices.visSectionParagraph];
            for (int row = 0; row < rowcount; row++)
            {
                fmt.Apply(update, (short) row);
            }
            update.Execute(shape);
        }


        public static void SetFormat(IVisio.Shape shape, VA.Text.ParagraphFormatCells format, int begin, int end)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            short rownum = -1;
            IVisio.Characters chars = null;

            int temp_leftindent = 0;
            int temp_indentfirst = 0;
            int temp_indentright = 0;
            int temp_spacingafter = 0;
            int temp_spacingbefore = 0;
            int temp_spacingline = 0;
            int temp_halign = 0;
            int temp_bulletindex = 0;
            int temp_bulletsize = 0;
            int temp_bulletfont = 0;

            VA.Text.TextHelper.SetRangeProps(shape, format.IndentLeft, IVisio.VisCellIndices.visIndentLeft,
                                             temp_leftindent, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.IndentFirst, IVisio.VisCellIndices.visIndentFirst,
                                             temp_indentfirst, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.IndentRight, IVisio.VisCellIndices.visIndentRight,
                                             temp_indentright, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.SpacingAfter, IVisio.VisCellIndices.visSpaceAfter,
                                             temp_spacingafter, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.SpacingBefore, IVisio.VisCellIndices.visSpaceBefore,
                                             temp_spacingbefore, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.SpacingLine, IVisio.VisCellIndices.visSpaceLine,
                                             temp_spacingline, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.HorizontalAlign, IVisio.VisCellIndices.visHorzAlign,
                                             temp_halign, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.BulletIndex, IVisio.VisCellIndices.visBulletIndex,
                                             temp_bulletindex, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.BulletSize, IVisio.VisCellIndices.visBulletFontSize,
                                             temp_bulletsize, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);
            VA.Text.TextHelper.SetRangeProps(shape, format.BulletFont, IVisio.VisCellIndices.visBulletFont,
                                             temp_bulletfont, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Paragraph);

            if (chars != null)
            {
                if (rownum < 0)
                {
                    throw new AutomationException("Internal Error");
                }

                var update = new VA.ShapeSheet.Update.SRCUpdate();
                format.Apply(update, rownum);
                update.Execute(shape);
            }
        }

        private static readonly VA.Text.ParagraphFormatQuery paraquery = new VA.Text.ParagraphFormatQuery();

        //TODO: Add Unit Test
        public static IList<ParagraphFormatCells> GetParagraphFormat(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var qds = paraquery.GetFormulasAndResults<double>(shape);

            var fmts = new List<ParagraphFormatCells>();
            for (int row = 0; row < qds.RowCount; row++)
            {
                var fmt = new ParagraphFormatCells();

                fmt.IndentFirst = qds.GetItem(row, paraquery.IndentFirst);
                fmt.IndentLeft = qds.GetItem(row, paraquery.IndentLeft);
                fmt.IndentRight = qds.GetItem(row, paraquery.IndentRight);
                fmt.SpacingAfter = qds.GetItem(row, paraquery.SpaceAfter);
                fmt.SpacingBefore = qds.GetItem(row, paraquery.SpaceBefore);
                fmt.SpacingLine = qds.GetItem(row, paraquery.SpaceLine);
                fmt.HorizontalAlign = qds.GetItem(row, paraquery.HorzAlign, v => (int) v);
                fmt.BulletIndex = qds.GetItem(row, paraquery.BulletIndex, v => (int) v);
                fmt.BulletFont = qds.GetItem(row, paraquery.BulletFont, v => (int) v);
                fmt.BulletSize = qds.GetItem(row, paraquery.BulletFontSize, v => (int) v);

                fmts.Add(fmt);
            }

            return fmts;
        }


        public static IList<VA.Text.CharacterFormatCells> GetCharacterFormat(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var qds = charquery.GetFormulasAndResults<double>(shape);

            var char_fmts = new List<CharacterFormatCells>();

            for (int row = 0; row < qds.RowCount; row++)
            {
                var fmt = new CharacterFormatCells();

                fmt.Color = qds.GetItem(row, charquery.Color, v => (int) v);
                fmt.Transparency = qds.GetItem(row, charquery.Trans);
                fmt.Font = qds.GetItem(row, charquery.Font, v => (int) v);
                fmt.Size = qds.GetItem(row, charquery.Size);
                fmt.Style = qds.GetItem(row, charquery.Style, v => (VA.Text.CharStyle) ((int) v));
                char_fmts.Add(fmt);
            }

            return char_fmts;
        }

        private const short char_section = (short) IVisio.VisSectionIndices.visSectionCharacter;

        public static void SetFormat(CharacterFormatCells format, IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            int rowcount = shape.RowCount[char_section];
            for (int row = 0; row < rowcount; row++)
            {
                format.Apply(update, (short) row);
            }

            update.Execute(shape);
        }

        public static void SetFormat(IVisio.Shape shape, VA.Text.CharacterFormatCells format, int begin, int end)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }


            short rownum = -1;
            IVisio.Characters chars = null;

            const int temp_color = 13;
            const int temp_size = 10;
            const int temp_font = 0;
            const int temp_style = 0;
            const int temp_trans = 0;

            VA.Text.TextHelper.SetRangeProps(shape, format.Color, IVisio.VisCellIndices.visCharacterColor, temp_color,
                                             begin, end, ref rownum, ref chars, VA.Text.TextHelper.rangetype.Character);
            VA.Text.TextHelper.SetRangeProps(shape, format.Size, IVisio.VisCellIndices.visCharacterSize, temp_size,
                                             begin, end, ref rownum, ref chars, VA.Text.TextHelper.rangetype.Character);
            VA.Text.TextHelper.SetRangeProps(shape, format.Font, IVisio.VisCellIndices.visCharacterFont, temp_font,
                                             begin, end, ref rownum, ref chars, VA.Text.TextHelper.rangetype.Character);
            VA.Text.TextHelper.SetRangeProps(shape, format.Style, IVisio.VisCellIndices.visCharacterStyle, temp_style,
                                             begin, end, ref rownum, ref chars, VA.Text.TextHelper.rangetype.Character);
            VA.Text.TextHelper.SetRangeProps(shape, format.Transparency, IVisio.VisCellIndices.visCharacterColorTrans,
                                             temp_trans, begin, end, ref rownum, ref chars,
                                             VA.Text.TextHelper.rangetype.Character);

            if (chars != null)
            {
                if (rownum < 0)
                {
                    throw new AutomationException("Internal Error");
                }

                var update = new VA.ShapeSheet.Update.SRCUpdate();
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Color.ForRow(rownum), format.Color.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Size.ForRow(rownum), format.Size.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Font.ForRow(rownum), format.Font.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Style.ForRow(rownum), format.Style.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_ColorTrans.ForRow(rownum),
                                            format.Transparency.Formula);
                update.Execute(shape);
            }
        }
    }
}
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
            VA.Text.TextHelper.SetRangeProps(shape, format.BulletFontSize, IVisio.VisCellIndices.visBulletFontSize,
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

        public static IList<List<ParagraphFormatCells>> GetParagraphFormat(IVisio.Page page, IList<int> shapeids)
        {
            return ParagraphFormatCells.GetCells(page,shapeids);
        }

        public static IList<ParagraphFormatCells> GetParagraphFormat(IVisio.Shape shape)
        {
            return ParagraphFormatCells.GetCells(shape);
        }

        public static IList<List<VA.Text.CharacterFormatCells>> GetCharacterFormat(IVisio.Page page, IList<int> shapeids)
        {
            return VA.Text.CharacterFormatCells.GetCells(page,shapeids);
        }

        public static IList<VA.Text.CharacterFormatCells> GetCharacterFormat(IVisio.Shape shape)
        {
            return VA.Text.CharacterFormatCells.GetCells(shape);
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

        private static readonly VA.ShapeSheet.SRC src_tabstopcount = VA.ShapeSheet.SRCConstants.Tabs_StopCount;
        private static readonly short unitcode_nocast = (short)IVisio.VisUnitCodes.visNoCast;
        private const short tab_section = (short)IVisio.VisSectionIndices.visSectionTab;

        public static void SetTabStops(IVisio.Shape shape, IList<TabStop> stops)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (stops == null)
            {
                throw new ArgumentNullException("stops");
            }

            ClearTabStops(shape);
            if (stops.Count < 1)
            {
                return;
            }

            const short row = 0;
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var vis_tab_stop_count = (short)IVisio.VisCellIndices.visTabStopCount;
            var tabstopcountcell = shape.CellsSRC[tab_section, row, vis_tab_stop_count];
            tabstopcountcell.FormulaU = stops.Count.ToString(invariant_culture);

            // set the number of tab stobs allowed for the shape
            var tagtab = GetTabTagForStops(stops.Count);
            shape.RowType[tab_section, (short)IVisio.VisRowIndices.visRowTab] = (short)tagtab;

            // add tab properties for each stop
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            for (int stop_index = 0; stop_index < stops.Count; stop_index++)
            {
                int i = stop_index * 3;

                var alignment = ((int)stops[stop_index].Alignment).ToString(invariant_culture);
                var position = ((int)stops[stop_index].Position).ToString(invariant_culture);

                var src_tabpos = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 1));
                var src_tabalign = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 2));
                var src_tabother = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 3));

                update.SetFormula(src_tabpos, position); // tab position
                update.SetFormula(src_tabalign, alignment); // tab alignment
                update.SetFormula(src_tabother, "0"); // tab unknown
            }

            update.Execute(shape);
        }

        public static IVisio.VisRowTags GetTabTagForStops(int stops)
        {
            if (stops < 0)
            {
                throw new ArgumentOutOfRangeException("stops");
            }

            var tagtab = IVisio.VisRowTags.visTagTab0;
            if ((0 <= stops) && (stops <= 2))
            {
                tagtab = IVisio.VisRowTags.visTagTab2;
            }
            else if ((3 <= stops) && (stops <= 10))
            {
                tagtab = IVisio.VisRowTags.visTagTab10;
            }
            else if ((11 <= stops) && (stops <= 60))
            {
                tagtab = IVisio.VisRowTags.visTagTab60;
            }
            else
            {
                throw new ArgumentOutOfRangeException("stops", "unsupported number of tabs");
            }

            return tagtab;
        }

        public static int GetTabStopCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            var tcell = shape.GetCell(src_tabstopcount);
            const short rounding = 0;
            return tcell.ResultInt[unitcode_nocast, rounding];
        }

        /// <summary>
        /// Remove all tab stops on the shape
        /// </summary>
        /// <param name="shape"></param>
        public static void ClearTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            int num_existing_tabstops = GetTabStopCount(shape);

            if (num_existing_tabstops < 1)
            {
                return;
            }

            var cell_tabstopcount = shape.GetCell(src_tabstopcount);
            cell_tabstopcount.FormulaForce = "0";

            const string formula = "0";

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            for (int i = 1; i < num_existing_tabstops * 3; i++)
            {
                var src = new VA.ShapeSheet.SRC(tab_section, (short)IVisio.VisRowIndices.visRowTab,
                                                (short)i);
                update.SetFormula(src, formula);
            }

            update.Execute(shape);
        }

        public static IVisio.Font TryGetFont(IVisio.Fonts fonts, string name)
        {
            try
            {
                var font = fonts[name];
                return font;
            }
            catch (System.Runtime.InteropServices.COMException comexc)
            {
                return null;
            }
        }
    }
}
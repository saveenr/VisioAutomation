using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using System;

namespace VisioAutomation.Text
{
    public class TextFormat
    {
        public IList<CharacterFormatCells> Character;
        public IList<ParagraphFormatCells> Paragraph;
        public TextBlockFormatCells TextBlock; 

        internal static IVisio.Characters SetRangeParagraphProps(IVisio.Shape shape, short cell, int value, int begin,
                                                         int end)
        {
            var chars = shape.Characters;
            chars.Begin = begin;
            chars.End = end;
            chars.ParaProps[cell] = (short)value;
            return chars;
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
                    chars.CharProps[(short)cell] = (short)value;
                    rownum = chars.CharPropsRow[(short)default_chars_bias];
                }
                else if (rt == rangetype.Paragraph)
                {
                    chars.ParaProps[(short)cell] = (short)value;
                    rownum = chars.ParaPropsRow[(short)default_chars_bias];
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
                run_begin = chars.RunBegin[(short)runtype];
                run_end = chars.RunEnd[(short)runtype];

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
            int rowcount = shape.RowCount[(short)IVisio.VisSectionIndices.visSectionCharacter];
            for (int row = 0; row < rowcount; row++)
            {
                fmt.Apply(update, (short)row);
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

            VA.Text.TextFormat.SetRangeProps(shape, format.IndentLeft, IVisio.VisCellIndices.visIndentLeft,
                                             temp_leftindent, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.IndentFirst, IVisio.VisCellIndices.visIndentFirst,
                                             temp_indentfirst, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.IndentRight, IVisio.VisCellIndices.visIndentRight,
                                             temp_indentright, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.SpacingAfter, IVisio.VisCellIndices.visSpaceAfter,
                                             temp_spacingafter, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.SpacingBefore, IVisio.VisCellIndices.visSpaceBefore,
                                             temp_spacingbefore, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.SpacingLine, IVisio.VisCellIndices.visSpaceLine,
                                             temp_spacingline, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.HorizontalAlign, IVisio.VisCellIndices.visHorzAlign,
                                             temp_halign, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.BulletIndex, IVisio.VisCellIndices.visBulletIndex,
                                             temp_bulletindex, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.BulletFontSize, IVisio.VisCellIndices.visBulletFontSize,
                                             temp_bulletsize, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);
            VA.Text.TextFormat.SetRangeProps(shape, format.BulletFont, IVisio.VisCellIndices.visBulletFont,
                                             temp_bulletfont, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Paragraph);

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

            VA.Text.TextFormat.SetRangeProps(shape, format.Color, IVisio.VisCellIndices.visCharacterColor, temp_color,
                                             begin, end, ref rownum, ref chars, VA.Text.TextFormat.rangetype.Character);
            VA.Text.TextFormat.SetRangeProps(shape, format.Size, IVisio.VisCellIndices.visCharacterSize, temp_size,
                                             begin, end, ref rownum, ref chars, VA.Text.TextFormat.rangetype.Character);
            VA.Text.TextFormat.SetRangeProps(shape, format.Font, IVisio.VisCellIndices.visCharacterFont, temp_font,
                                             begin, end, ref rownum, ref chars, VA.Text.TextFormat.rangetype.Character);
            VA.Text.TextFormat.SetRangeProps(shape, format.Style, IVisio.VisCellIndices.visCharacterStyle, temp_style,
                                             begin, end, ref rownum, ref chars, VA.Text.TextFormat.rangetype.Character);
            VA.Text.TextFormat.SetRangeProps(shape, format.Transparency, IVisio.VisCellIndices.visCharacterColorTrans,
                                             temp_trans, begin, end, ref rownum, ref chars,
                                             VA.Text.TextFormat.rangetype.Character);

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

        public static TextFormat GetFormat(IVisio.Shape shape)
        {
            var t = new TextFormat();
            t.Character = VA.Text.CharacterFormatCells.GetCells(shape);
            t.Paragraph = VA.Text.ParagraphFormatCells.GetCells(shape);
            t.TextBlock = VA.Text.TextBlockFormatCells.GetCells(shape);
            return t;
        }

        public static IList<TextFormat> GetFormat(IVisio.Page page, IList<int> shapeids)
        {
            var c = VA.Text.CharacterFormatCells.GetCells(page, shapeids);
            var p = VA.Text.ParagraphFormatCells.GetCells(page, shapeids);
            var b = VA.Text.TextBlockFormatCells.GetCells(page, shapeids);

            var l = new List<TextFormat>(shapeids.Count);
            for (int i = 0; i < shapeids.Count; i++)
            {
                var t = new TextFormat();
                t.Character = c[i];
                t.Paragraph = p[i];
                t.TextBlock = b[i];
                l.Add(t);
                
            }
            return l;
        }

        public static VA.Text.CharStyle GetCharStyle(bool bold, bool italic, bool underline, bool smallcaps)
        {
            VA.Text.CharStyle cs = 0;
            if (bold)
            {
                cs |= VA.Text.CharStyle.Bold;
            }

            if (italic)
            {
                cs |= VA.Text.CharStyle.Italic;
            }

            if (underline)
            {
                cs |= VA.Text.CharStyle.UnderLine;
            }

            if (smallcaps)
            {
                cs |= VA.Text.CharStyle.SmallCaps;
            }

            return cs;
        }

    }
}
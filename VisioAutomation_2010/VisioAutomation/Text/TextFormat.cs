using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using System;
using System.Linq;

namespace VisioAutomation.Text
{
    public class TextFormat
    {
        private IList<CharacterFormatCells> _characterFormats;
        private IList<ParagraphFormatCells> _paragraphFormats;
        private TextBlockFormatCells _textBlocks;
        private IList<TextRun> _characterTextRuns;
        private IList<TextRun> _paragraphTextRuns;
        private IList<TabStop> _tabStops;

        public IList<CharacterFormatCells> CharacterFormats
        {
            get { return _characterFormats; }
        }

        public IList<ParagraphFormatCells> ParagraphFormats
        {
            get { return _paragraphFormats; }
        }

        public TextBlockFormatCells TextBlocks
        {
            get { return _textBlocks; }
        }

        public IList<TextRun> CharacterTextRuns
        {
            get { return _characterTextRuns; }
        }

        public IList<TextRun> ParagraphTextRuns
        {
            get { return _paragraphTextRuns; }
        }

        public IList<TabStop> TabStops
        {
            get { return _tabStops; }
        }

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


        private static IList<TextRun> GetTextRuns(
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

        private static readonly VA.ShapeSheet.SRC src_tabstopcount = VA.ShapeSheet.SRCConstants.Tabs_StopCount;
        private static readonly short unitcode_nocast = (short)IVisio.VisUnitCodes.visNoCast;
        private const short tab_section = (short)IVisio.VisSectionIndices.visSectionTab;

        private static IList<TabStop> GetTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            int num_stops = VA.Text.TextFormat.GetTabStopCount(shape);

            if (num_stops < 1)
            {
                return new List<TabStop>(0);
            }

            const short row = 0;


            var cellsrcs = new List<VA.ShapeSheet.SRC>(num_stops*3);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                int i = stop_index * 3;


                var src_tabpos = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 1));
                var src_tabalign = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 2));
                var src_tabother = new VA.ShapeSheet.SRC(tab_section, row, (short)(i + 3));

                cellsrcs.Add(src_tabpos);
                cellsrcs.Add(src_tabalign );
                cellsrcs.Add(src_tabother);
            }

            var stream = VA.ShapeSheet.SRC.ToStream(cellsrcs);
            var formulas = VA.ShapeSheet.Query.QueryUtil.GetFormulasU(shape, stream, num_stops*3);
            var unitcodes = cellsrcs.Select(i => IVisio.VisUnitCodes.visNoCast).ToList();
            var results = VA.ShapeSheet.Query.QueryUtil.GetResults<double>(shape, stream, unitcodes, num_stops*3);

            var stops_list = new List<TabStop>(num_stops);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                var pos = results[(stop_index*3) + 1];
                var align = (VA.Text.TabStopAlignment) ((int)results[(stop_index*3) + 2]);
                var ts = new VA.Text.TabStop(pos, align);
                stops_list.Add(ts);
            }

            return stops_list;
        }

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

        private static IVisio.VisRowTags GetTabTagForStops(int stops)
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

        private static int GetTabStopCount(IVisio.Shape shape)
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
        private static void ClearTabStops(IVisio.Shape shape)
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
            t._characterFormats = VA.Text.CharacterFormatCells.GetCells(shape);
            t._paragraphFormats = VA.Text.ParagraphFormatCells.GetCells(shape);
            t._textBlocks = VA.Text.TextBlockFormatCells.GetCells(shape);
            t._characterTextRuns = VA.Text.TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visCharPropRow, true);
            t._paragraphTextRuns = VA.Text.TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visParaPropRow, true);
            t._tabStops = VA.Text.TextFormat.GetTabStops(shape);
            return t;
        }

        public static IList<TextFormat> GetFormat(IVisio.Page page, IList<int> shapeids)
        {
            var c = VA.Text.CharacterFormatCells.GetCells(page, shapeids);
            var p = VA.Text.ParagraphFormatCells.GetCells(page, shapeids);
            var b = VA.Text.TextBlockFormatCells.GetCells(page, shapeids);
            var page_shapes = page.Shapes;
            var l = new List<TextFormat>(shapeids.Count);
            for (int i = 0; i < shapeids.Count; i++)
            {
                var t = new TextFormat();
                t._characterFormats = c[i];
                t._paragraphFormats = p[i];
                t._textBlocks = b[i];
                l.Add(t);

                var shape = page_shapes.get_ItemFromID(shapeids[i]);
                t._characterTextRuns = VA.Text.TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visCharPropRow, true);
                t._paragraphTextRuns = VA.Text.TextFormat.GetTextRuns(shape, IVisio.VisRunTypes.visParaPropRow, true);

                t._tabStops = VA.Text.TextFormat.GetTabStops(shape);
            }

            return l;
        }
    }
}
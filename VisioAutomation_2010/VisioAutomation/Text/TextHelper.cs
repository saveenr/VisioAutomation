using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using System.Linq;

namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        private const short tab_section = (short)IVisio.VisSectionIndices.visSectionTab;

        public static List<TabStop> GetTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_stops = GetTabStopCount(shape);

            if (num_stops < 1)
            {
                return new List<TabStop>(0);
            }

            const short row = 0;
            
            var stream = new VisioAutomation.ShapeSheet.Streams.SrcStreamBuilder(num_stops * 3);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                int i = stop_index * 3;

                var src_tabpos = new ShapeSheet.Src(tab_section, row, (short)(i + 1));
                var src_tabalign = new ShapeSheet.Src(tab_section, row, (short)(i + 2));
                var src_tabother = new ShapeSheet.Src(tab_section, row, (short)(i + 3));

                stream.Add(src_tabpos);
                stream.Add(src_tabalign);
                stream.Add(src_tabother);
            }

            var surface = new SurfaceTarget(shape);

            const object[] unitcodes = null;

            var results = surface.GetResults<double>(stream.ToStream(), unitcodes);

            var stops_list = new List<TabStop>(num_stops);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                var pos = results[(stop_index * 3) + 1];
                var align = (TabStopAlignment)((int)results[(stop_index * 3) + 2]);
                var ts = new TabStop(pos, align);
                stops_list.Add(ts);
            }

            return stops_list;
        }

        public static void SetTabStops(IVisio.Shape shape, IList<TabStop> stops)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (stops == null)
            {
                throw new System.ArgumentNullException(nameof(stops));
            }

            ClearTabStops(shape);
            if (stops.Count < 1)
            {
                return;
            }

            const short row = 0;
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var vis_tab_stop_count = (short)IVisio.VisCellIndices.visTabStopCount;
            var tabstopcountcell = shape.CellsSRC[tab_section, row, vis_tab_stop_count];
            tabstopcountcell.FormulaU = stops.Count.ToString(culture);

            // set the number of tab stobs allowed for the shape
            var tagtab = GetTabTagForStops(stops.Count);
            shape.RowType[tab_section, (short)IVisio.VisRowIndices.visRowTab] = (short)tagtab;

            // add tab properties for each stop
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            for (int stop_index = 0; stop_index < stops.Count; stop_index++)
            {
                int i = stop_index * 3;

                var alignment = ((int)stops[stop_index].Alignment).ToString(culture);
                var position = ((int)stops[stop_index].Position).ToString(culture);

                var src_tabpos = new ShapeSheet.Src(tab_section, row, (short)(i + 1));
                var src_tabalign = new ShapeSheet.Src(tab_section, row, (short)(i + 2));
                var src_tabother = new ShapeSheet.Src(tab_section, row, (short)(i + 3));

                writer.SetFormula(src_tabpos, position); // tab position
                writer.SetFormula(src_tabalign, alignment); // tab alignment
                writer.SetFormula(src_tabother, "0"); // tab unknown
            }

            writer.Commit(shape);
        }

        private static IVisio.VisRowTags GetTabTagForStops(int stops)
        {
            if (stops < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(stops));
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
                throw new System.ArgumentOutOfRangeException(nameof(stops), "unsupported number of tabs");
            }

            return tagtab;
        }

        private static int GetTabStopCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var cell_tabstopcount = shape.CellsSRC[ShapeSheet.SrcConstants.TabStopCount.Section, ShapeSheet.SrcConstants.TabStopCount.Row, ShapeSheet.SrcConstants.TabStopCount.Cell];
            const short rounding = 0;

            return cell_tabstopcount.ResultInt[(short)IVisio.VisUnitCodes.visNumber, rounding];
        }

        /// <summary>
        /// Remove all tab stops on the shape
        /// </summary>
        /// <param name="shape"></param>
        private static void ClearTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_existing_tabstops = GetTabStopCount(shape);

            if (num_existing_tabstops < 1)
            {
                return;
            }

            var cell_tabstopcount = shape.CellsSRC[ShapeSheet.SrcConstants.TabStopCount.Section, ShapeSheet.SrcConstants.TabStopCount.Row, ShapeSheet.SrcConstants.TabStopCount.Cell];
            cell_tabstopcount.FormulaForce = "0";

            const string formula = "0";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            for (int i = 1; i < num_existing_tabstops * 3; i++)
            {
                var src = new ShapeSheet.Src(tab_section, (short)IVisio.VisRowIndices.visRowTab,
                    (short)i);
                writer.SetFormula(src, formula);
            }

            writer.Commit(shape);
        }

        public static List<List<CharacterFormatCells>> GetCharacterFormatCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<CharacterFormatCells> GetCharacterFormatCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = CharacterFormatCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> CharacterFormatCells_lazy_reader = new System.Lazy<CharacterFormatCellsReader>();


        class CharacterFormatCellsReader : CellGroupReader<Text.CharacterFormatCells>
        {
            public CharacterFormatCellsReader() :
                base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                InitializeQuery();
            }

            public override Text.CharacterFormatCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.CharacterFormatCells();

                var cols = this.query_multirow.SectionQueries[0].Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.Color = getcellvalue(nameof(CharacterFormatCells.Color));
                cells.ColorTransparency = getcellvalue(nameof(CharacterFormatCells.ColorTransparency));
                cells.Font = getcellvalue(nameof(CharacterFormatCells.Font));
                cells.Size = getcellvalue(nameof(CharacterFormatCells.Size));
                cells.Style = getcellvalue(nameof(CharacterFormatCells.Style));
                cells.AsianFont = getcellvalue(nameof(CharacterFormatCells.AsianFont));
                cells.AsianFont = getcellvalue(nameof(CharacterFormatCells.AsianFont));
                cells.Case = getcellvalue(nameof(CharacterFormatCells.Case));
                cells.ComplexScriptFont = getcellvalue(nameof(CharacterFormatCells.ComplexScriptFont));
                cells.ComplexScriptSize = getcellvalue(nameof(CharacterFormatCells.ComplexScriptSize));
                cells.DoubleStrikethrough = getcellvalue(nameof(CharacterFormatCells.DoubleStrikethrough));
                cells.DoubleUnderline = getcellvalue(nameof(CharacterFormatCells.DoubleUnderline));
                cells.FontScale = getcellvalue(nameof(CharacterFormatCells.FontScale));
                cells.LangID = getcellvalue(nameof(CharacterFormatCells.LangID));
                cells.Letterspace = getcellvalue(nameof(CharacterFormatCells.Letterspace));
                cells.Locale = getcellvalue(nameof(CharacterFormatCells.Locale));
                cells.LocalizeFont = getcellvalue(nameof(CharacterFormatCells.LocalizeFont));
                cells.Overline = getcellvalue(nameof(CharacterFormatCells.Overline));
                cells.Perpendicular = getcellvalue(nameof(CharacterFormatCells.Perpendicular));
                cells.Pos = getcellvalue(nameof(CharacterFormatCells.Pos));
                cells.RTLText = getcellvalue(nameof(CharacterFormatCells.RTLText));
                cells.Strikethru = getcellvalue(nameof(CharacterFormatCells.Strikethru));
                cells.UseVertical = getcellvalue(nameof(CharacterFormatCells.UseVertical));

                return cells;
            }
        }

        public static List<List<ParagraphFormatCells>> GetParagraphFormatCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ParagraphFormatCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ParagraphFormatCells> GetParagraphFormatCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ParagraphFormatCells_lazy_reader.Value;
            return reader.GetCellsMultiRow(shape, type);
        }


        private static readonly System.Lazy<ParagraphFormatCellsReader> ParagraphFormatCells_lazy_reader = new System.Lazy<ParagraphFormatCellsReader>();


        class ParagraphFormatCellsReader : CellGroupReader<Text.ParagraphFormatCells>
        {
            public ParagraphFormatCellsReader() : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                InitializeQuery();
            }

            public override Text.ParagraphFormatCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.ParagraphFormatCells();


                var cols = this.query_multirow.SectionQueries[0].Columns;
                var names = cells.CellMetadata.Select(i => i.Name).ToList();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }



                cells.IndentFirst = getcellvalue(nameof(ParagraphFormatCells.IndentFirst));
                cells.IndentLeft = getcellvalue(nameof(ParagraphFormatCells.IndentLeft));
                cells.IndentRight = getcellvalue(nameof(ParagraphFormatCells.IndentRight));
                cells.SpacingAfter = getcellvalue(nameof(ParagraphFormatCells.SpacingAfter));
                cells.SpacingBefore = getcellvalue(nameof(ParagraphFormatCells.SpacingBefore));
                cells.SpacingLine = getcellvalue(nameof(ParagraphFormatCells.SpacingLine));
                cells.HorizontalAlign = getcellvalue(nameof(ParagraphFormatCells.HorizontalAlign));
                cells.Bullet = getcellvalue(nameof(ParagraphFormatCells.Bullet));
                cells.BulletFont = getcellvalue(nameof(ParagraphFormatCells.BulletFont));
                cells.BulletFontSize = getcellvalue(nameof(ParagraphFormatCells.BulletFontSize));
                cells.LocalizeBulletFont = getcellvalue(nameof(ParagraphFormatCells.LocalizeBulletFont));
                cells.TextPosAfterBullet = getcellvalue(nameof(ParagraphFormatCells.TextPosAfterBullet));
                cells.Flags = getcellvalue(nameof(ParagraphFormatCells.Flags));
                cells.BulletString = getcellvalue(nameof(ParagraphFormatCells.BulletString));

                return cells;
            }
        }

        public static IList<TextBlockCells> GetTextBlockCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = TextBlockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextBlockCells GetTextBlockCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = TextBlockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<TextBlockCellsReader> TextBlockCells_lazy_reader = new System.Lazy<TextBlockCellsReader>();

        class TextBlockCellsReader : CellGroupReader<Text.TextBlockCells>
        {

            public TextBlockCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                InitializeQuery();
            }

            public override Text.TextBlockCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextBlockCells();
                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.BottomMargin = getcellvalue(nameof(TextBlockCells.BottomMargin));
                cells.LeftMargin = getcellvalue(nameof(TextBlockCells.LeftMargin));
                cells.RightMargin = getcellvalue(nameof(TextBlockCells.RightMargin));
                cells.TopMargin = getcellvalue(nameof(TextBlockCells.TopMargin));
                cells.DefaultTabStop = getcellvalue(nameof(TextBlockCells.DefaultTabStop));
                cells.Background = getcellvalue(nameof(TextBlockCells.Background));
                cells.BackgroundTransparency = getcellvalue(nameof(TextBlockCells.BackgroundTransparency));
                cells.Direction = getcellvalue(nameof(TextBlockCells.Direction));
                cells.VerticalAlign = getcellvalue(nameof(TextBlockCells.VerticalAlign));

                return cells;
            }
        }

        public static List<TextXFormCells> GetTextXFormCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = TextXFormCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextXFormCells GetTextXFormCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = TextXFormCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<TextXFormCellsReader> TextXFormCells_lazy_reader = new System.Lazy<TextXFormCellsReader>();


        class TextXFormCellsReader : CellGroupReader<Text.TextXFormCells>
        {
            public TextXFormCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                InitializeQuery();
            }

            public override Text.TextXFormCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextXFormCells();

                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.PinX = getcellvalue(nameof(TextXFormCells.PinX));
                cells.PinY = getcellvalue(nameof(TextXFormCells.PinY));
                cells.LocPinX = getcellvalue(nameof(TextXFormCells.LocPinX));
                cells.LocPinY = getcellvalue(nameof(TextXFormCells.LocPinY));
                cells.Width = getcellvalue(nameof(TextXFormCells.Width));
                cells.Height = getcellvalue(nameof(TextXFormCells.Height));
                cells.Angle = getcellvalue(nameof(TextXFormCells.Angle));

                return cells;
            }
        }

    }
}
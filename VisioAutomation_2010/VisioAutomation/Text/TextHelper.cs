using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

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
            public SectionQueryColumn Font { get; set; }
            public SectionQueryColumn Style { get; set; }
            public SectionQueryColumn Color { get; set; }
            public SectionQueryColumn Size { get; set; }
            public SectionQueryColumn ColorTransparency { get; set; }
            public SectionQueryColumn AsianFont { get; set; }
            public SectionQueryColumn Case { get; set; }
            public SectionQueryColumn ComplexScriptFont { get; set; }
            public SectionQueryColumn ComplexScriptSize { get; set; }
            public SectionQueryColumn DoubleStrikethrough { get; set; }
            public SectionQueryColumn DoubleUnderline { get; set; }
            public SectionQueryColumn LangID { get; set; }
            public SectionQueryColumn Locale { get; set; }
            public SectionQueryColumn LocalizeFont { get; set; }
            public SectionQueryColumn Overline { get; set; }
            public SectionQueryColumn Perpendicular { get; set; }
            public SectionQueryColumn Pos { get; set; }
            public SectionQueryColumn RTLText { get; set; }
            public SectionQueryColumn FontScale { get; set; }
            public SectionQueryColumn Letterspace { get; set; }
            public SectionQueryColumn Strikethru { get; set; }
            public SectionQueryColumn UseVertical { get; set; }

            public CharacterFormatCellsReader() :
                base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionCharacter);

                this.Color = sec.Columns.Add(SrcConstants.CharColor, nameof(this.Color));
                this.ColorTransparency = sec.Columns.Add(SrcConstants.CharColorTransparency, nameof(this.ColorTransparency));
                this.Font = sec.Columns.Add(SrcConstants.CharFont, nameof(this.Font));
                this.Size = sec.Columns.Add(SrcConstants.CharSize, nameof(this.Size));
                this.Style = sec.Columns.Add(SrcConstants.CharStyle, nameof(this.Style));
                this.AsianFont = sec.Columns.Add(SrcConstants.CharAsianFont, nameof(this.AsianFont));
                this.Case = sec.Columns.Add(SrcConstants.CharCase, nameof(this.Case));
                this.ComplexScriptFont = sec.Columns.Add(SrcConstants.CharComplexScriptFont, nameof(this.ComplexScriptFont));
                this.ComplexScriptSize = sec.Columns.Add(SrcConstants.CharComplexScriptSize, nameof(this.ComplexScriptSize));
                this.DoubleStrikethrough = sec.Columns.Add(SrcConstants.CharDoubleStrikethrough, nameof(this.DoubleStrikethrough));
                this.DoubleUnderline = sec.Columns.Add(SrcConstants.CharDoubleUnderline, nameof(this.DoubleUnderline));
                this.LangID = sec.Columns.Add(SrcConstants.CharLangID, nameof(this.LangID));
                this.Locale = sec.Columns.Add(SrcConstants.CharLocale, nameof(this.Locale));
                this.LocalizeFont = sec.Columns.Add(SrcConstants.CharLocalizeFont, nameof(this.LocalizeFont));
                this.Overline = sec.Columns.Add(SrcConstants.CharOverline, nameof(this.Overline));
                this.Perpendicular = sec.Columns.Add(SrcConstants.CharPerpendicular, nameof(this.Perpendicular));
                this.Pos = sec.Columns.Add(SrcConstants.CharPos, nameof(this.Pos));
                this.RTLText = sec.Columns.Add(SrcConstants.CharRTLText, nameof(this.RTLText));
                this.FontScale = sec.Columns.Add(SrcConstants.CharFontScale, nameof(this.FontScale));
                this.Letterspace = sec.Columns.Add(SrcConstants.CharLetterspace, nameof(this.Letterspace));
                this.Strikethru = sec.Columns.Add(SrcConstants.CharStrikethru, nameof(this.Strikethru));
                this.UseVertical = sec.Columns.Add(SrcConstants.CharUseVertical, nameof(this.UseVertical));

            }

            public override Text.CharacterFormatCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.CharacterFormatCells();
                cells.Color = row[this.Color];
                cells.ColorTransparency = row[this.ColorTransparency];
                cells.Font = row[this.Font];
                cells.Size = row[this.Size];
                cells.Style = row[this.Style];
                cells.AsianFont = row[this.AsianFont];
                cells.AsianFont = row[this.AsianFont];
                cells.Case = row[this.Case];
                cells.ComplexScriptFont = row[this.ComplexScriptFont];
                cells.ComplexScriptSize = row[this.ComplexScriptSize];
                cells.DoubleStrikethrough = row[this.DoubleStrikethrough];
                cells.DoubleUnderline = row[this.DoubleUnderline];
                cells.FontScale = row[this.FontScale];
                cells.LangID = row[this.LangID];
                cells.Letterspace = row[this.Letterspace];
                cells.Locale = row[this.Locale];
                cells.LocalizeFont = row[this.LocalizeFont];
                cells.Overline = row[this.Overline];
                cells.Perpendicular = row[this.Perpendicular];
                cells.Pos = row[this.Pos];
                cells.RTLText = row[this.RTLText];
                cells.Strikethru = row[this.Strikethru];
                cells.UseVertical = row[this.UseVertical];

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
            public SectionQueryColumn Bullet { get; set; }
            public SectionQueryColumn BulletFont { get; set; }
            public SectionQueryColumn BulletFontSize { get; set; }
            public SectionQueryColumn BulletString { get; set; }
            public SectionQueryColumn Flags { get; set; }
            public SectionQueryColumn HorizontalAlign { get; set; }
            public SectionQueryColumn IndentFirst { get; set; }
            public SectionQueryColumn IndentLeft { get; set; }
            public SectionQueryColumn IndentRight { get; set; }
            public SectionQueryColumn LocalizeBulletFont { get; set; }
            public SectionQueryColumn SpaceAfter { get; set; }
            public SectionQueryColumn SpaceBefore { get; set; }
            public SectionQueryColumn SpaceLine { get; set; }
            public SectionQueryColumn TextPosAfterBullet { get; set; }

            public ParagraphFormatCellsReader() : base(new VisioAutomation.ShapeSheet.Query.SectionsQuery())
            {
                var sec = this.query_multirow.SectionQueries.Add(IVisio.VisSectionIndices.visSectionParagraph);
                this.Bullet = sec.Columns.Add(SrcConstants.ParaBullet, nameof(this.Bullet));
                this.BulletFont = sec.Columns.Add(SrcConstants.ParaBulletFont, nameof(this.BulletFont));
                this.BulletFontSize = sec.Columns.Add(SrcConstants.ParaBulletFontSize, nameof(this.BulletFontSize));
                this.BulletString = sec.Columns.Add(SrcConstants.ParaBulletString, nameof(this.BulletString));
                this.Flags = sec.Columns.Add(SrcConstants.ParaFlags, nameof(this.Flags));
                this.HorizontalAlign = sec.Columns.Add(SrcConstants.ParaHorizontalAlign, nameof(this.HorizontalAlign));
                this.IndentFirst = sec.Columns.Add(SrcConstants.ParaIndentFirst, nameof(this.IndentFirst));
                this.IndentLeft = sec.Columns.Add(SrcConstants.ParaIndentLeft, nameof(this.IndentLeft));
                this.IndentRight = sec.Columns.Add(SrcConstants.ParaIndentRight, nameof(this.IndentRight));
                this.LocalizeBulletFont = sec.Columns.Add(SrcConstants.ParaLocalizeBulletFont, nameof(this.LocalizeBulletFont));
                this.SpaceAfter = sec.Columns.Add(SrcConstants.ParaSpacingAfter, nameof(this.SpaceAfter));
                this.SpaceBefore = sec.Columns.Add(SrcConstants.ParaSpacingBefore, nameof(this.SpaceBefore));
                this.SpaceLine = sec.Columns.Add(SrcConstants.ParaSpacingLine, nameof(this.SpaceLine));
                this.TextPosAfterBullet = sec.Columns.Add(SrcConstants.ParaTextPosAfterBullet, nameof(this.TextPosAfterBullet));
            }

            public override Text.ParagraphFormatCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.ParagraphFormatCells();
                cells.IndentFirst = row[this.IndentFirst];
                cells.IndentLeft = row[this.IndentLeft];
                cells.IndentRight = row[this.IndentRight];
                cells.SpacingAfter = row[this.SpaceAfter];
                cells.SpacingBefore = row[this.SpaceBefore];
                cells.SpacingLine = row[this.SpaceLine];
                cells.HorizontalAlign = row[this.HorizontalAlign];
                cells.Bullet = row[this.Bullet];
                cells.BulletFont = row[this.BulletFont];
                cells.BulletFontSize = row[this.BulletFontSize];
                cells.LocalizeBulletFont = row[this.LocalizeBulletFont];
                cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
                cells.Flags = row[this.Flags];
                cells.BulletString = row[this.BulletString];

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
            public CellColumn BottomMargin { get; set; }
            public CellColumn LeftMargin { get; set; }
            public CellColumn RightMargin { get; set; }
            public CellColumn TopMargin { get; set; }
            public CellColumn DefaultTabStop { get; set; }
            public CellColumn Background { get; set; }
            public CellColumn BackgroundTransparency { get; set; }
            public CellColumn Direction { get; set; }
            public CellColumn VerticalAlign { get; set; }

            public TextBlockCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.BottomMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBottomMargin, nameof(this.BottomMargin));
                this.LeftMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockLeftMargin, nameof(this.LeftMargin));
                this.RightMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockRightMargin, nameof(this.RightMargin));
                this.TopMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockTopMargin, nameof(this.TopMargin));
                this.DefaultTabStop = this.query_singlerow.Columns.Add(SrcConstants.TextBlockDefaultTabStop, nameof(this.DefaultTabStop));
                this.Background = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBackground, nameof(this.Background));
                this.BackgroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBackgroundTransparency, nameof(this.BackgroundTransparency));
                this.Direction = this.query_singlerow.Columns.Add(SrcConstants.TextBlockDirection, nameof(this.Direction));
                this.VerticalAlign = this.query_singlerow.Columns.Add(SrcConstants.TextBlockVerticalAlign, nameof(this.VerticalAlign));

            }

            public override Text.TextBlockCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextBlockCells();
                cells.BottomMargin = row[this.BottomMargin];
                cells.LeftMargin = row[this.LeftMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.TopMargin = row[this.TopMargin];
                cells.DefaultTabStop = row[this.DefaultTabStop];
                cells.Background = row[this.Background];
                cells.BackgroundTransparency = row[this.BackgroundTransparency];
                cells.Direction = row[this.Direction];
                cells.VerticalAlign = row[this.VerticalAlign];
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
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public TextXFormCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.PinX = this.query_singlerow.Columns.Add(SrcConstants.TextXFormPinX, nameof(this.PinX));
                this.PinY = this.query_singlerow.Columns.Add(SrcConstants.TextXFormPinY, nameof(this.PinY));
                this.LocPinX = this.query_singlerow.Columns.Add(SrcConstants.TextXFormLocPinX, nameof(this.LocPinX));
                this.LocPinY = this.query_singlerow.Columns.Add(SrcConstants.TextXFormLocPinY, nameof(this.LocPinY));
                this.Width = this.query_singlerow.Columns.Add(SrcConstants.TextXFormWidth, nameof(this.Width));
                this.Height = this.query_singlerow.Columns.Add(SrcConstants.TextXFormHeight, nameof(this.Height));
                this.Angle = this.query_singlerow.Columns.Add(SrcConstants.TextXFormAngle, nameof(this.Angle));

            }

            public override Text.TextXFormCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextXFormCells();
                cells.PinX = row[this.PinX];
                cells.PinY = row[this.PinY];
                cells.LocPinX = row[this.LocPinX];
                cells.LocPinY = row[this.LocPinY];
                cells.Width = row[this.Width];
                cells.Height = row[this.Height];
                cells.Angle = row[this.Angle];
                return cells;
            }
        }

    }
}
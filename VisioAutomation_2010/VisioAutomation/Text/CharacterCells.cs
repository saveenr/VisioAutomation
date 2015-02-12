using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class CharacterCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public VA.ShapeSheet.CellData<int> Color { get; set; }
        public VA.ShapeSheet.CellData<int> Font { get; set; }
        public VA.ShapeSheet.CellData<double> Size { get; set; }
        public VA.ShapeSheet.CellData<int> Style { get; set; }
        public VA.ShapeSheet.CellData<double> Transparency { get; set; }
        public VA.ShapeSheet.CellData<int> AsianFont { get; set; }
        public VA.ShapeSheet.CellData<int> Case { get; set; }
        public VA.ShapeSheet.CellData<int> ComplexScriptFont { get; set; }
        public VA.ShapeSheet.CellData<double> ComplexScriptSize { get; set; }
        public VA.ShapeSheet.CellData<bool> DoubleStrikeThrough { get; set; }
        public VA.ShapeSheet.CellData<bool> DoubleUnderline { get; set; }
        public VA.ShapeSheet.CellData<int> LangID { get; set; }
        public VA.ShapeSheet.CellData<int> Locale { get; set; }
        public VA.ShapeSheet.CellData<int> LocalizeFont { get; set; }
        public VA.ShapeSheet.CellData<bool> Overline { get; set; }
        public VA.ShapeSheet.CellData<bool> Perpendicular { get; set; }
        public VA.ShapeSheet.CellData<int> Pos { get; set; }
        public VA.ShapeSheet.CellData<int> RTLText { get; set; }
        public VA.ShapeSheet.CellData<double> FontScale { get; set; }
        public VA.ShapeSheet.CellData<double> Letterspace { get; set; }
        public VA.ShapeSheet.CellData<bool> Strikethru { get; set; }
        public VA.ShapeSheet.CellData<int> UseVertical { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(VA.ShapeSheet.SRCConstants.CharColor, this.Color.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharFont, this.Font.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharSize, this.Size.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharStyle, this.Style.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharColorTrans, this.Transparency.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharAsianFont, this.AsianFont.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharCase, this.Case.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharComplexScriptFont, this.ComplexScriptFont.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharComplexScriptSize, this.ComplexScriptSize.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharDblUnderline, this.DoubleUnderline.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough, this.DoubleStrikeThrough.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharLangID, this.LangID.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharFontScale, this.FontScale.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharLangID, this.LangID.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharLetterspace, this.Letterspace.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharLocale, this.Locale.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharLocalizeFont, this.LocalizeFont.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharOverline, this.Overline.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharPerpendicular, this.Perpendicular.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharPos, this.Pos.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharRTLText, this.RTLText.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharStrikethru, this.Strikethru.Formula);
                yield return newpair(VA.ShapeSheet.SRCConstants.CharUseVertical, this.UseVertical.Formula);
            }
        }

        public static IList<List<CharacterCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells<CharacterCells,double>(page, shapeids, query, query.GetCells);
        }

        public static IList<CharacterCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<CharacterCells,double>(shape, query, query.GetCells);
        }

        private static CharacterFormatCellQuery _mCellQuery;

        private static CharacterFormatCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new CharacterFormatCellQuery();
            return _mCellQuery;
        }

        class CharacterFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn Font { get; set; }
            public CellColumn Style { get; set; }
            public CellColumn Color { get; set; }
            public CellColumn Size { get; set; }
            public CellColumn Trans { get; set; }
            public CellColumn AsianFont { get; set; }
            public CellColumn Case { get; set; }
            public CellColumn ComplexScriptFont { get; set; }
            public CellColumn ComplexScriptSize { get; set; }
            public CellColumn DoubleStrikethrough { get; set; }
            public CellColumn DoubleUnderline { get; set; }
            public CellColumn LangID { get; set; }
            public CellColumn Locale { get; set; }
            public CellColumn LocalizeFont { get; set; }
            public CellColumn Overline { get; set; }
            public CellColumn Perpendicular { get; set; }
            public CellColumn Pos { get; set; }
            public CellColumn RTLText { get; set; }
            public CellColumn FontScale { get; set; }
            public CellColumn Letterspace { get; set; }
            public CellColumn Strikethru { get; set; }
            public CellColumn UseVertical { get; set; }

            public CharacterFormatCellQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);
                Color = sec.AddCell(VA.ShapeSheet.SRCConstants.CharColor, "Color");
                Trans = sec.AddCell(VA.ShapeSheet.SRCConstants.CharColorTrans, "Trans");
                Font = sec.AddCell(VA.ShapeSheet.SRCConstants.CharFont, "Font");
                Size = sec.AddCell(VA.ShapeSheet.SRCConstants.CharSize, "Size");
                Style = sec.AddCell(VA.ShapeSheet.SRCConstants.CharStyle, "Style");
                AsianFont = sec.AddCell(VA.ShapeSheet.SRCConstants.CharAsianFont, "AsianFont");
                Case = sec.AddCell(VA.ShapeSheet.SRCConstants.CharCase, "Case");
                ComplexScriptFont = sec.AddCell(VA.ShapeSheet.SRCConstants.CharComplexScriptFont, "ComplexScriptStyle");
                ComplexScriptSize = sec.AddCell(VA.ShapeSheet.SRCConstants.CharComplexScriptSize, "ComplexScriptSize");
                DoubleStrikethrough = sec.AddCell(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough, "DoubleStrikethrough");
                DoubleUnderline = sec.AddCell(VA.ShapeSheet.SRCConstants.CharDblUnderline, "DoubleUnderline");
                LangID = sec.AddCell(VA.ShapeSheet.SRCConstants.CharLangID, "LangID");
                Locale = sec.AddCell(VA.ShapeSheet.SRCConstants.CharLocale, "Locale");
                LocalizeFont = sec.AddCell(VA.ShapeSheet.SRCConstants.CharLocalizeFont, "LocalizeFont");
                Overline = sec.AddCell(VA.ShapeSheet.SRCConstants.CharOverline, "Overline");
                Perpendicular = sec.AddCell(VA.ShapeSheet.SRCConstants.CharPerpendicular, "Perpendicular");
                Pos = sec.AddCell(VA.ShapeSheet.SRCConstants.CharPos, "Pos");
                RTLText = sec.AddCell(VA.ShapeSheet.SRCConstants.CharRTLText, "RTLText");
                FontScale = sec.AddCell(VA.ShapeSheet.SRCConstants.CharFontScale, "FontScale");
                Letterspace = sec.AddCell(VA.ShapeSheet.SRCConstants.CharLetterspace, "Letterspace");
                Strikethru = sec.AddCell(VA.ShapeSheet.SRCConstants.CharStrikethru, "Strikethru");
                UseVertical = sec.AddCell(VA.ShapeSheet.SRCConstants.CharUseVertical, "UseVertical");
            }

            public CharacterCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new CharacterCells();
                cells.Color = row[this.Color.Ordinal].ToInt();
                cells.Transparency = row[this.Trans.Ordinal];
                cells.Font = row[this.Font.Ordinal].ToInt();
                cells.Size = row[this.Size.Ordinal];
                cells.Style = row[this.Style.Ordinal].ToInt();
                cells.AsianFont = row[this.AsianFont.Ordinal].ToInt();
                cells.AsianFont = row[this.AsianFont.Ordinal].ToInt();
                cells.Case = row[this.Case.Ordinal].ToInt();
                cells.ComplexScriptFont = row[this.ComplexScriptFont.Ordinal].ToInt();
                cells.ComplexScriptSize = row[this.ComplexScriptSize.Ordinal];
                cells.DoubleStrikeThrough = row[this.DoubleStrikethrough.Ordinal].ToBool();
                cells.DoubleUnderline = row[this.DoubleUnderline.Ordinal].ToBool();
                cells.FontScale = row[this.FontScale.Ordinal];
                cells.LangID = row[this.LangID.Ordinal].ToInt();
                cells.Letterspace = row[this.Letterspace.Ordinal];
                cells.Locale = row[this.Locale.Ordinal].ToInt();
                cells.LocalizeFont = row[this.LocalizeFont.Ordinal].ToInt();
                cells.Overline = row[this.Overline.Ordinal].ToBool();
                cells.Perpendicular = row[this.Perpendicular.Ordinal].ToBool();
                cells.Pos = row[this.Pos.Ordinal].ToInt();
                cells.RTLText = row[this.RTLText.Ordinal].ToInt();
                cells.Strikethru = row[this.Strikethru.Ordinal].ToBool();
                cells.UseVertical = row[this.UseVertical.Ordinal].ToInt();

                return cells;
            }
        }
    }
}
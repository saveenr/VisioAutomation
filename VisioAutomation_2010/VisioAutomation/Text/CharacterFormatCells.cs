using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : VA.ShapeSheet.CellGroups.CellGroupMultiRow
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

        public override void ApplyFormulasForRow(ApplyFormula func, short row)
        {
            func(VA.ShapeSheet.SRCConstants.CharColor.ForRow(row), this.Color.Formula);
            func(VA.ShapeSheet.SRCConstants.CharFont.ForRow(row), this.Font.Formula);
            func(VA.ShapeSheet.SRCConstants.CharSize.ForRow(row), this.Size.Formula);
            func(VA.ShapeSheet.SRCConstants.CharStyle.ForRow(row), this.Style.Formula);
            func(VA.ShapeSheet.SRCConstants.CharColorTrans.ForRow(row), this.Transparency.Formula);
            
            func(VA.ShapeSheet.SRCConstants.CharAsianFont.ForRow(row), this.AsianFont.Formula);
            func(VA.ShapeSheet.SRCConstants.CharCase.ForRow(row), this.Case.Formula);
            func(VA.ShapeSheet.SRCConstants.CharComplexScriptFont.ForRow(row), this.ComplexScriptFont.Formula);
            func(VA.ShapeSheet.SRCConstants.CharComplexScriptSize.ForRow(row), this.ComplexScriptSize.Formula);
            
            func(VA.ShapeSheet.SRCConstants.CharDblUnderline.ForRow(row), this.DoubleUnderline.Formula);
            func(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough.ForRow(row), this.DoubleStrikeThrough.Formula);
            func(VA.ShapeSheet.SRCConstants.CharLangID.ForRow(row), this.LangID.Formula);

            func(VA.ShapeSheet.SRCConstants.CharFontScale.ForRow(row), this.FontScale.Formula);
            func(VA.ShapeSheet.SRCConstants.CharLangID.ForRow(row), this.LangID.Formula);
            func(VA.ShapeSheet.SRCConstants.CharLetterspace.ForRow(row), this.Letterspace.Formula);
            func(VA.ShapeSheet.SRCConstants.CharLocale.ForRow(row), this.Locale.Formula);

            func(VA.ShapeSheet.SRCConstants.CharLocalizeFont.ForRow(row), this.LocalizeFont.Formula);
            func(VA.ShapeSheet.SRCConstants.CharOverline.ForRow(row), this.Overline.Formula);
            
            func(VA.ShapeSheet.SRCConstants.CharPerpendicular.ForRow(row), this.Perpendicular.Formula);
            func(VA.ShapeSheet.SRCConstants.CharPos.ForRow(row), this.Pos.Formula);

            func(VA.ShapeSheet.SRCConstants.CharRTLText.ForRow(row), this.RTLText.Formula);
            func(VA.ShapeSheet.SRCConstants.CharStrikethru.ForRow(row), this.Strikethru.Formula);
            func(VA.ShapeSheet.SRCConstants.CharUseVertical.ForRow(row), this.UseVertical.Formula);


        }

        public static IList<List<CharacterFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return _GetCells(page, shapeids, query, query.GetCells);
        }

        public static IList<CharacterFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells(shape, query, query.GetCells);
        }

        private static CharacterFormatCellQuery _mCellQuery;

        private static CharacterFormatCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new CharacterFormatCellQuery();
            return _mCellQuery;
        }

        class CharacterFormatCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.CellQuery.Column Font { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Style { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Color { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Size { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Trans { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column AsianFont { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Case { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column ComplexScriptFont { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column ComplexScriptSize { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column DoubleStrikethrough { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column DoubleUnderline { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column LangID { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Locale { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column LocalizeFont { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Overline { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Perpendicular { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Pos { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column RTLText { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column FontScale { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Letterspace { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column Strikethru { get; set; }
            public VA.ShapeSheet.Query.CellQuery.Column UseVertical { get; set; }

            public CharacterFormatCellQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);
                Color = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharColor, "Color");
                Trans = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharColorTrans, "Trans");
                Font = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharFont, "Font");
                Size = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharSize, "Size");
                Style = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharStyle, "Style");
                AsianFont = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharAsianFont, "AsianFont");
                Case = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharCase, "Case");
                ComplexScriptFont = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharComplexScriptFont, "ComplexScriptStyle");
                ComplexScriptSize = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharComplexScriptSize, "ComplexScriptSize");
                DoubleStrikethrough = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough, "DoubleStrikethrough");
                DoubleUnderline = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharDblUnderline, "DoubleUnderline");
                LangID = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharLangID, "LangID");
                Locale = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharLocale, "Locale");
                LocalizeFont = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharLocalizeFont, "LocalizeFont");
                Overline = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharOverline, "Overline");
                Perpendicular = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharPerpendicular, "Perpendicular");
                Pos = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharPos, "Pos");
                RTLText = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharRTLText, "RTLText");
                FontScale = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharFontScale, "FontScale");
                Letterspace = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharLetterspace, "Letterspace");
                Strikethru = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharStrikethru, "Strikethru");
                UseVertical = sec.Columns.Add(VA.ShapeSheet.SRCConstants.CharUseVertical, "UseVertical");
            }

            public CharacterFormatCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new CharacterFormatCells();
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
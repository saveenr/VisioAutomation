using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : VA.ShapeSheet.CellGroups.CellGroupMultiRowEx
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

        private static CharacterFormatQuery m_query;

        private static CharacterFormatQuery get_query()
        {
            m_query = m_query ?? new CharacterFormatQuery();
            return m_query;
        }

        class CharacterFormatQuery : VA.ShapeSheet.Query.QueryEx
        {
            public int Font { get; set; }
            public int Style { get; set; }
            public int Color { get; set; }
            public int Size { get; set; }
            public int Trans { get; set; }
            public int AsianFont { get; set; }
            public int Case { get; set; }
            public int ComplexScriptFont { get; set; }
            public int ComplexScriptSize { get; set; }
            public int DoubleStrikethrough { get; set; }
            public int DoubleUnderline { get; set; }
            public int LangID { get; set; }
            public int Locale { get; set; }
            public int LocalizeFont { get; set; }
            public int Overline { get; set; }
            public int Perpendicular { get; set; }
            public int Pos { get; set; }
            public int RTLText { get; set; }
            public int FontScale { get; set; }
            public int Letterspace { get; set; }
            public int Strikethru { get; set; }
            public int UseVertical { get; set; }

            public CharacterFormatQuery() 
            {
                var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);
                Color = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharColor, "Color");
                Trans = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharColorTrans, "Trans");
                Font = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharFont, "Font");
                Size = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharSize, "Size");
                Style = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharStyle, "Style");
                AsianFont = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharAsianFont, "AsianFont");
                Case = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharCase, "Case");
                ComplexScriptFont= sec.AddColumn(VA.ShapeSheet.SRCConstants.CharComplexScriptFont, "Style");
                ComplexScriptSize= sec.AddColumn(VA.ShapeSheet.SRCConstants.CharComplexScriptSize, "Style");
                DoubleStrikethrough = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough, "DoubleStrikethrough");
                DoubleUnderline = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharDblUnderline, "DoubleUnderline");
                LangID = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharLangID, "LangID");
                Locale = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharLocale, "Locale");
                LocalizeFont = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharLocalizeFont, "LocalizeFont");
                Overline = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharOverline, "Overline");
                Perpendicular = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharPerpendicular, "Perpendicular");
                Pos = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharPos, "Pos");
                RTLText = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharRTLText, "RTLText");
                FontScale = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharFontScale, "FontScale");
                Letterspace = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharLetterspace, "Letterspace");
                Strikethru = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharStrikethru, "Strikethru");
                UseVertical = sec.AddColumn(VA.ShapeSheet.SRCConstants.CharUseVertical, "UseVertical");
            }

            public CharacterFormatCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
                var cells = new CharacterFormatCells();
                cells.Color = row[this.Color].ToInt();
                cells.Transparency = row[this.Trans];
                cells.Font = row[this.Font].ToInt();
                cells.Size = row[this.Size];
                cells.Style = row[this.Style].ToInt();
                cells.AsianFont = row[this.AsianFont].ToInt();
                cells.AsianFont = row[this.AsianFont].ToInt();
                cells.Case = row[this.Case].ToInt();
                cells.ComplexScriptFont = row[this.ComplexScriptFont].ToInt();
                cells.ComplexScriptSize = row[this.ComplexScriptSize];
                cells.DoubleStrikeThrough = row[this.DoubleStrikethrough].ToBool();
                cells.DoubleUnderline = row[this.DoubleUnderline].ToBool();
                cells.FontScale = row[this.FontScale];
                cells.LangID = row[this.LangID].ToInt();
                cells.Letterspace = row[this.Letterspace];
                cells.Locale = row[this.Locale].ToInt();
                cells.LocalizeFont = row[this.LocalizeFont].ToInt();
                cells.Overline = row[this.Overline].ToBool();
                cells.Perpendicular = row[this.Perpendicular].ToBool();
                cells.Pos = row[this.Pos].ToInt();
                cells.RTLText = row[this.RTLText].ToInt();
                cells.Strikethru = row[this.Strikethru].ToBool();
                cells.UseVertical = row[this.UseVertical].ToInt();

                return cells;
            }
        }
    }
}
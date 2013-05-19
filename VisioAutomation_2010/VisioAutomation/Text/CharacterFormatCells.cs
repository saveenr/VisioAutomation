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
            var query = new CharacterFormatQuery();
            return VA.ShapeSheet.CellGroups.CellGroupMultiRow.CellsFromRowsGrouped(page, shapeids, query, get_cells_from_row);
        }

        public static IList<CharacterFormatCells> GetCells(IVisio.Shape shape)
        {
            var query = new CharacterFormatQuery();
            return VA.ShapeSheet.CellGroups.CellGroupMultiRow.CellsFromRows(shape, query, get_cells_from_row);
        }

        private static CharacterFormatCells get_cells_from_row(CharacterFormatQuery query, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new CharacterFormatCells();
            cells.Color = table[row,query.Color].ToInt();
            cells.Transparency = table[row,query.Trans];
            cells.Font = table[row,query.Font].ToInt();
            cells.Size = table[row,query.Size];
            cells.Style = table[row,query.Style].ToInt();
            cells.AsianFont = table[row,query.AsianFont].ToInt();
            cells.AsianFont = table[row,query.AsianFont].ToInt();
            cells.Case = table[row,query.Case].ToInt();
            cells.ComplexScriptFont= table[row,query.ComplexScriptFont].ToInt();
            cells.ComplexScriptSize= table[row,query.ComplexScriptSize];
            cells.DoubleStrikeThrough= table[row,query.DoubleStrikethrough].ToBool();
            cells.DoubleUnderline= table[row,query.DoubleUnderline].ToBool();
            cells.FontScale= table[row,query.FontScale];
            cells.LangID= table[row,query.LangID].ToInt();
            cells.Letterspace= table[row,query.Letterspace];
            cells.Locale= table[row,query.Locale].ToInt();
            cells.LocalizeFont= table[row,query.LocalizeFont].ToInt();
            cells.Overline= table[row,query.Overline].ToBool();
            cells.Perpendicular= table[row,query.Perpendicular].ToBool();          
            cells.Pos= table[row,query.Pos].ToInt();
            cells.RTLText= table[row,query.RTLText].ToInt();
            cells.Strikethru= table[row,query.Strikethru].ToBool();
            cells.UseVertical= table[row,query.UseVertical].ToInt();

            return cells;
        }

        class CharacterFormatQuery : VA.ShapeSheet.Query.SectionQuery
        {
            public VA.ShapeSheet.Query.QueryColumn Font { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Style { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Color { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Size { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Trans { get; set; }
            public VA.ShapeSheet.Query.QueryColumn AsianFont { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Case { get; set; }
            public VA.ShapeSheet.Query.QueryColumn ComplexScriptFont { get; set; }
            public VA.ShapeSheet.Query.QueryColumn ComplexScriptSize { get; set; }
            public VA.ShapeSheet.Query.QueryColumn DoubleStrikethrough { get; set; }
            public VA.ShapeSheet.Query.QueryColumn DoubleUnderline { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LangID { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Locale { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LocalizeFont { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Overline { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Perpendicular { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Pos { get; set; }
            public VA.ShapeSheet.Query.QueryColumn RTLText { get; set; }
            public VA.ShapeSheet.Query.QueryColumn FontScale { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Letterspace { get; set; }
            public VA.ShapeSheet.Query.QueryColumn Strikethru { get; set; }
            public VA.ShapeSheet.Query.QueryColumn UseVertical { get; set; }

            public CharacterFormatQuery() :
                base(IVisio.VisSectionIndices.visSectionCharacter)
            {
                Color = this.AddColumn(VA.ShapeSheet.SRCConstants.CharColor, "Color");
                Trans = this.AddColumn(VA.ShapeSheet.SRCConstants.CharColorTrans, "Trans");
                Font = this.AddColumn(VA.ShapeSheet.SRCConstants.CharFont, "Font");
                Size = this.AddColumn(VA.ShapeSheet.SRCConstants.CharSize, "Size");
                Style = this.AddColumn(VA.ShapeSheet.SRCConstants.CharStyle, "Style");
                AsianFont = this.AddColumn(VA.ShapeSheet.SRCConstants.CharAsianFont, "AsianFont");
                Case = this.AddColumn(VA.ShapeSheet.SRCConstants.CharCase, "Case");
                ComplexScriptFont= this.AddColumn(VA.ShapeSheet.SRCConstants.CharComplexScriptFont, "Style");
                ComplexScriptSize= this.AddColumn(VA.ShapeSheet.SRCConstants.CharComplexScriptSize, "Style");
                DoubleStrikethrough = this.AddColumn(VA.ShapeSheet.SRCConstants.CharDoubleStrikethrough, "DoubleStrikethrough");
                DoubleUnderline = this.AddColumn(VA.ShapeSheet.SRCConstants.CharDblUnderline, "DoubleUnderline");
                LangID = this.AddColumn(VA.ShapeSheet.SRCConstants.CharLangID, "LangID");
                Locale = this.AddColumn(VA.ShapeSheet.SRCConstants.CharLocale, "Locale");
                LocalizeFont = this.AddColumn(VA.ShapeSheet.SRCConstants.CharLocalizeFont, "LocalizeFont");
                Overline = this.AddColumn(VA.ShapeSheet.SRCConstants.CharOverline, "Overline");
                Perpendicular = this.AddColumn(VA.ShapeSheet.SRCConstants.CharPerpendicular, "Perpendicular");
                Pos = this.AddColumn(VA.ShapeSheet.SRCConstants.CharPos, "Pos");
                RTLText = this.AddColumn(VA.ShapeSheet.SRCConstants.CharRTLText, "RTLText");
                FontScale = this.AddColumn(VA.ShapeSheet.SRCConstants.CharFontScale, "FontScale");
                Letterspace = this.AddColumn(VA.ShapeSheet.SRCConstants.CharLetterspace, "Letterspace");
                Strikethru = this.AddColumn(VA.ShapeSheet.SRCConstants.CharStrikethru, "Strikethru");
                UseVertical = this.AddColumn(VA.ShapeSheet.SRCConstants.CharUseVertical, "UseVertical");
            }
        }
    }
}
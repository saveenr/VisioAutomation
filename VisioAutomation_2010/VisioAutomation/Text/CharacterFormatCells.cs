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
            func(VA.ShapeSheet.SRCConstants.Char_Color.ForRow(row), this.Color.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Font.ForRow(row), this.Font.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Size.ForRow(row), this.Size.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Style.ForRow(row), this.Style.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_ColorTrans.ForRow(row), this.Transparency.Formula);
            
            func(VA.ShapeSheet.SRCConstants.Char_AsianFont.ForRow(row), this.AsianFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Case.ForRow(row), this.Case.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_ComplexScriptFont.ForRow(row), this.ComplexScriptFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_ComplexScriptSize.ForRow(row), this.ComplexScriptSize.Formula);
            
            func(VA.ShapeSheet.SRCConstants.Char_DblUnderline.ForRow(row), this.DoubleUnderline.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_DoubleStrikethrough.ForRow(row), this.DoubleStrikeThrough.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_LangID.ForRow(row), this.LangID.Formula);

            func(VA.ShapeSheet.SRCConstants.Char_FontScale.ForRow(row), this.FontScale.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_LangID.ForRow(row), this.LangID.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Letterspace.ForRow(row), this.Letterspace.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Locale.ForRow(row), this.Locale.Formula);

            func(VA.ShapeSheet.SRCConstants.Char_LocalizeFont.ForRow(row), this.LocalizeFont.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Overline.ForRow(row), this.Overline.Formula);
            
            func(VA.ShapeSheet.SRCConstants.Char_Perpendicular.ForRow(row), this.Perpendicular.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Pos.ForRow(row), this.Pos.Formula);

            func(VA.ShapeSheet.SRCConstants.Char_RTLText.ForRow(row), this.RTLText.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_Strikethru.ForRow(row), this.Strikethru.Formula);
            func(VA.ShapeSheet.SRCConstants.Char_UseVertical.ForRow(row), this.UseVertical.Formula);


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
                Color = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Color, "Color");
                Trans = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ColorTrans, "Trans");
                Font = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Font, "Font");
                Size = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size, "Size");
                Style = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Style, "Style");
                AsianFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_AsianFont, "AsianFont");
                Case = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Case, "Case");
                ComplexScriptFont= this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ComplexScriptFont, "Style");
                ComplexScriptSize= this.AddColumn(VA.ShapeSheet.SRCConstants.Char_ComplexScriptSize, "Style");
                DoubleStrikethrough = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_DoubleStrikethrough, "DoubleStrikethrough");
                DoubleUnderline = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_DblUnderline, "DoubleUnderline");
                LangID = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_LangID, "LangID");
                Locale = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Locale, "Locale");
                LocalizeFont = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_LocalizeFont, "LocalizeFont");
                Overline = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Overline, "Overline");
                Perpendicular = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Perpendicular, "Perpendicular");
                Pos = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Pos, "Pos");
                RTLText = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_RTLText, "RTLText");
                FontScale = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_FontScale, "FontScale");
                Letterspace = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Letterspace, "Letterspace");
                Strikethru = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_Strikethru, "Strikethru");
                UseVertical = this.AddColumn(VA.ShapeSheet.SRCConstants.Char_UseVertical, "UseVertical");
            }
        }
    }
}
namespace VisioAutomation.ShapeSheet.Query.Common
{
    class CharacterFormatCellsQuery : CellQuery
    {
        public Query.CellColumn Font { get; set; }
        public Query.CellColumn Style { get; set; }
        public Query.CellColumn Color { get; set; }
        public Query.CellColumn Size { get; set; }
        public Query.CellColumn Trans { get; set; }
        public Query.CellColumn AsianFont { get; set; }
        public Query.CellColumn Case { get; set; }
        public Query.CellColumn ComplexScriptFont { get; set; }
        public Query.CellColumn ComplexScriptSize { get; set; }
        public Query.CellColumn DoubleStrikethrough { get; set; }
        public Query.CellColumn DoubleUnderline { get; set; }
        public Query.CellColumn LangID { get; set; }
        public Query.CellColumn Locale { get; set; }
        public Query.CellColumn LocalizeFont { get; set; }
        public Query.CellColumn Overline { get; set; }
        public Query.CellColumn Perpendicular { get; set; }
        public Query.CellColumn Pos { get; set; }
        public Query.CellColumn RTLText { get; set; }
        public Query.CellColumn FontScale { get; set; }
        public Query.CellColumn Letterspace { get; set; }
        public Query.CellColumn Strikethru { get; set; }
        public Query.CellColumn UseVertical { get; set; }

        public CharacterFormatCellsQuery()
        {
            var sec = this.AddSection(Microsoft.Office.Interop.Visio.VisSectionIndices.visSectionCharacter);





            this.Color = sec.AddCell(ShapeSheet.SRCConstants.CharColor, nameof(ShapeSheet.SRCConstants.CharColor));
            this.Trans = sec.AddCell(ShapeSheet.SRCConstants.CharColorTrans, nameof(ShapeSheet.SRCConstants.CharColorTrans));
            this.Font = sec.AddCell(ShapeSheet.SRCConstants.CharFont, nameof(ShapeSheet.SRCConstants.CharFont));
            this.Size = sec.AddCell(ShapeSheet.SRCConstants.CharSize, nameof(ShapeSheet.SRCConstants.CharSize));
            this.Style = sec.AddCell(ShapeSheet.SRCConstants.CharStyle, nameof(ShapeSheet.SRCConstants.CharStyle));
            this.AsianFont = sec.AddCell(ShapeSheet.SRCConstants.CharAsianFont, nameof(ShapeSheet.SRCConstants.CharAsianFont));
            this.Case = sec.AddCell(ShapeSheet.SRCConstants.CharCase, nameof(ShapeSheet.SRCConstants.CharCase));
            this.ComplexScriptFont = sec.AddCell(ShapeSheet.SRCConstants.CharComplexScriptFont, nameof(ShapeSheet.SRCConstants.CharComplexScriptFont));
            this.ComplexScriptSize = sec.AddCell(ShapeSheet.SRCConstants.CharComplexScriptSize, nameof(ShapeSheet.SRCConstants.CharComplexScriptSize));
            this.DoubleStrikethrough = sec.AddCell(ShapeSheet.SRCConstants.CharDoubleStrikethrough, nameof(ShapeSheet.SRCConstants.CharDoubleStrikethrough));
            this.DoubleUnderline = sec.AddCell(ShapeSheet.SRCConstants.CharDblUnderline, nameof(ShapeSheet.SRCConstants.CharDblUnderline));
            this.LangID = sec.AddCell(ShapeSheet.SRCConstants.CharLangID, nameof(ShapeSheet.SRCConstants.CharLangID));
            this.Locale = sec.AddCell(ShapeSheet.SRCConstants.CharLocale, nameof(ShapeSheet.SRCConstants.CharLocale));
            this.LocalizeFont = sec.AddCell(ShapeSheet.SRCConstants.CharLocalizeFont, nameof(ShapeSheet.SRCConstants.CharLocalizeFont));
            this.Overline = sec.AddCell(ShapeSheet.SRCConstants.CharOverline, nameof(ShapeSheet.SRCConstants.CharOverline));
            this.Perpendicular = sec.AddCell(ShapeSheet.SRCConstants.CharPerpendicular, nameof(ShapeSheet.SRCConstants.CharPerpendicular));
            this.Pos = sec.AddCell(ShapeSheet.SRCConstants.CharPos, nameof(ShapeSheet.SRCConstants.CharPos));
            this.RTLText = sec.AddCell(ShapeSheet.SRCConstants.CharRTLText, nameof(ShapeSheet.SRCConstants.CharRTLText));
            this.FontScale = sec.AddCell(ShapeSheet.SRCConstants.CharFontScale, nameof(ShapeSheet.SRCConstants.CharFontScale));
            this.Letterspace = sec.AddCell(ShapeSheet.SRCConstants.CharLetterspace, nameof(ShapeSheet.SRCConstants.CharLetterspace));
            this.Strikethru = sec.AddCell(ShapeSheet.SRCConstants.CharStrikethru, nameof(ShapeSheet.SRCConstants.CharStrikethru));
            this.UseVertical = sec.AddCell(ShapeSheet.SRCConstants.CharUseVertical, nameof(ShapeSheet.SRCConstants.CharUseVertical));

        }

        public VisioAutomation.Text.CharacterCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Text.CharacterCells();
            cells.Color = Extensions.CellDataMethods.ToInt(row[this.Color]);
            cells.Transparency = row[this.Trans];
            cells.Font = Extensions.CellDataMethods.ToInt(row[this.Font]);
            cells.Size = row[this.Size];
            cells.Style = Extensions.CellDataMethods.ToInt(row[this.Style]);
            cells.AsianFont = Extensions.CellDataMethods.ToInt(row[this.AsianFont]);
            cells.AsianFont = Extensions.CellDataMethods.ToInt(row[this.AsianFont]);
            cells.Case = Extensions.CellDataMethods.ToInt(row[this.Case]);
            cells.ComplexScriptFont = Extensions.CellDataMethods.ToInt(row[this.ComplexScriptFont]);
            cells.ComplexScriptSize = row[this.ComplexScriptSize];
            cells.DoubleStrikeThrough = Extensions.CellDataMethods.ToBool(row[this.DoubleStrikethrough]);
            cells.DoubleUnderline = Extensions.CellDataMethods.ToBool(row[this.DoubleUnderline]);
            cells.FontScale = row[this.FontScale];
            cells.LangID = Extensions.CellDataMethods.ToInt(row[this.LangID]);
            cells.Letterspace = row[this.Letterspace];
            cells.Locale = Extensions.CellDataMethods.ToInt(row[this.Locale]);
            cells.LocalizeFont = Extensions.CellDataMethods.ToInt(row[this.LocalizeFont]);
            cells.Overline = Extensions.CellDataMethods.ToBool(row[this.Overline]);
            cells.Perpendicular = Extensions.CellDataMethods.ToBool(row[this.Perpendicular]);
            cells.Pos = Extensions.CellDataMethods.ToInt(row[this.Pos]);
            cells.RTLText = Extensions.CellDataMethods.ToInt(row[this.RTLText]);
            cells.Strikethru = Extensions.CellDataMethods.ToBool(row[this.Strikethru]);
            cells.UseVertical = Extensions.CellDataMethods.ToInt(row[this.UseVertical]);

            return cells;
        }
    }
}
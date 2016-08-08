using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class CharacterFormatCellsQuery : CellQuery
    {
        public VisioAutomation.ShapeSheetQuery.CellColumn Font { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Style { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Color { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Size { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Trans { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn AsianFont { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Case { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn ComplexScriptFont { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn ComplexScriptSize { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn DoubleStrikethrough { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn DoubleUnderline { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn LangID { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Locale { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn LocalizeFont { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Overline { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Perpendicular { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Pos { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn RTLText { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn FontScale { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Letterspace { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn Strikethru { get; set; }
        public VisioAutomation.ShapeSheetQuery.CellColumn UseVertical { get; set; }

        public CharacterFormatCellsQuery()
        {
            var sec = this.AddSection(IVisio.VisSectionIndices.visSectionCharacter);

            this.Color = sec.AddCell(SRCCON.CharColor, nameof(SRCCON.CharColor));
            this.Trans = sec.AddCell(SRCCON.CharColorTrans, nameof(SRCCON.CharColorTrans));
            this.Font = sec.AddCell(SRCCON.CharFont, nameof(SRCCON.CharFont));
            this.Size = sec.AddCell(SRCCON.CharSize, nameof(SRCCON.CharSize));
            this.Style = sec.AddCell(SRCCON.CharStyle, nameof(SRCCON.CharStyle));
            this.AsianFont = sec.AddCell(SRCCON.CharAsianFont, nameof(SRCCON.CharAsianFont));
            this.Case = sec.AddCell(SRCCON.CharCase, nameof(SRCCON.CharCase));
            this.ComplexScriptFont = sec.AddCell(SRCCON.CharComplexScriptFont, nameof(SRCCON.CharComplexScriptFont));
            this.ComplexScriptSize = sec.AddCell(SRCCON.CharComplexScriptSize, nameof(SRCCON.CharComplexScriptSize));
            this.DoubleStrikethrough = sec.AddCell(SRCCON.CharDoubleStrikethrough, nameof(SRCCON.CharDoubleStrikethrough));
            this.DoubleUnderline = sec.AddCell(SRCCON.CharDblUnderline, nameof(SRCCON.CharDblUnderline));
            this.LangID = sec.AddCell(SRCCON.CharLangID, nameof(SRCCON.CharLangID));
            this.Locale = sec.AddCell(SRCCON.CharLocale, nameof(SRCCON.CharLocale));
            this.LocalizeFont = sec.AddCell(SRCCON.CharLocalizeFont, nameof(SRCCON.CharLocalizeFont));
            this.Overline = sec.AddCell(SRCCON.CharOverline, nameof(SRCCON.CharOverline));
            this.Perpendicular = sec.AddCell(SRCCON.CharPerpendicular, nameof(SRCCON.CharPerpendicular));
            this.Pos = sec.AddCell(SRCCON.CharPos, nameof(SRCCON.CharPos));
            this.RTLText = sec.AddCell(SRCCON.CharRTLText, nameof(SRCCON.CharRTLText));
            this.FontScale = sec.AddCell(SRCCON.CharFontScale, nameof(SRCCON.CharFontScale));
            this.Letterspace = sec.AddCell(SRCCON.CharLetterspace, nameof(SRCCON.CharLetterspace));
            this.Strikethru = sec.AddCell(SRCCON.CharStrikethru, nameof(SRCCON.CharStrikethru));
            this.UseVertical = sec.AddCell(SRCCON.CharUseVertical, nameof(SRCCON.CharUseVertical));

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
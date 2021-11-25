using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Text
{
    public class CharacterFormatting
    {
        public Core.CellValue AsianFont { get; set; }
        public Core.CellValue Case { get; set; }
        public Core.CellValue Color { get; set; }
        public Core.CellValue ComplexScriptFont { get; set; }
        public Core.CellValue ComplexScriptSize { get; set; }
        public Core.CellValue DoubleStrikeThrough { get; set; }
        public Core.CellValue DoubleUnderline { get; set; }
        public Core.CellValue Font { get; set; }
        public Core.CellValue FontScale { get; set; }
        public Core.CellValue LangID { get; set; }
        public Core.CellValue Letterspace { get; set; }
        public Core.CellValue Locale { get; set; }
        public Core.CellValue LocalizeFont { get; set; }
        public Core.CellValue Overline { get; set; }
        public Core.CellValue Perpendicular { get; set; }
        public Core.CellValue Pos { get; set; }
        public Core.CellValue RTLText { get; set; }
        public Core.CellValue Size { get; set; }
        public Core.CellValue Strikethru { get; set; }
        public Core.CellValue Style { get; set; }
        public Core.CellValue Transparency { get; set; }
        public Core.CellValue UseVertical { get; set; }

        internal void ApplyFormulas(VisioAutomation.ShapeSheet.Writers.SrcWriter writer, short row)
        {
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharColor.CloneWithNewRow(row), this.Color);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharFont.CloneWithNewRow(row), this.Font);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharSize.CloneWithNewRow(row), this.Size);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharStyle.CloneWithNewRow(row), this.Style);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharColorTransparency.CloneWithNewRow(row), this.Transparency);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharAsianFont.CloneWithNewRow(row), this.AsianFont);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharCase.CloneWithNewRow(row), this.Case);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharComplexScriptFont.CloneWithNewRow(row), this.ComplexScriptFont);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharComplexScriptSize.CloneWithNewRow(row), this.ComplexScriptSize);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharDoubleUnderline.CloneWithNewRow(row), this.DoubleUnderline);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharDoubleStrikethrough.CloneWithNewRow(row), this.DoubleStrikeThrough);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharFontScale.CloneWithNewRow(row), this.FontScale);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharLetterspace.CloneWithNewRow(row), this.Letterspace);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharLocale.CloneWithNewRow(row), this.Locale);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharLocalizeFont.CloneWithNewRow(row), this.LocalizeFont);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharOverline.CloneWithNewRow(row), this.Overline);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharPerpendicular.CloneWithNewRow(row), this.Perpendicular);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharPos.CloneWithNewRow(row), this.Pos);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharRTLText.CloneWithNewRow(row), this.RTLText);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharStrikethru.CloneWithNewRow(row), this.Strikethru);
            writer.SetValue(VisioAutomation.Core.SrcConstants.CharUseVertical.CloneWithNewRow(row), this.UseVertical);
        }

        public void ApplyFormulasTo(CharacterFormatting target)
        {
            if (this.AsianFont.HasValue) { target.AsianFont = this.AsianFont; }
            if (this.Case.HasValue) { target.Case = this.Case; }
            if (this.Color.HasValue) { target.Color = this.Color; }
            if (this.ComplexScriptFont.HasValue) { target.ComplexScriptFont = this.ComplexScriptFont; }
            if (this.ComplexScriptSize.HasValue) { target.ComplexScriptSize = this.ComplexScriptSize; }
            if (this.DoubleStrikeThrough.HasValue) { target.DoubleStrikeThrough = this.DoubleStrikeThrough; }
            if (this.DoubleUnderline.HasValue) { target.DoubleUnderline = this.DoubleUnderline; }
            if (this.Font.HasValue) { target.Font = this.Font; }
            if (this.LangID.HasValue) { target.LangID = this.LangID; }
            if (this.Locale.HasValue) { target.Locale = this.Locale; }
            if (this.LocalizeFont.HasValue) { target.LocalizeFont = this.LocalizeFont; }
            if (this.Overline.HasValue) { target.Overline = this.Overline; }
            if (this.Perpendicular.HasValue) { target.Perpendicular = this.Perpendicular; }
            if (this.Pos.HasValue) { target.Pos = this.Pos; }
            if (this.RTLText.HasValue) { target.RTLText = this.RTLText; }
            if (this.FontScale.HasValue) { target.FontScale = this.FontScale; }
            if (this.Size.HasValue) { target.Size = this.Size; }
            if (this.Letterspace.HasValue) { target.Letterspace = this.Letterspace; }
            if (this.Strikethru.HasValue) { target.Strikethru = this.Strikethru; }
            if (this.Style.HasValue) { target.Style = this.Style; }
            if (this.Transparency.HasValue) { target.Transparency = this.Transparency; }
            if (this.UseVertical.HasValue) { target.UseVertical = this.UseVertical; }
        }
    }
}
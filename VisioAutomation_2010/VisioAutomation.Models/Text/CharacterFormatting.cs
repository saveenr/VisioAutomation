using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Text
{
    public class CharacterFormatting
    {
        public ShapeSheet.CellValueLiteral AsianFont { get; set; }
        public ShapeSheet.CellValueLiteral Case { get; set; }
        public ShapeSheet.CellValueLiteral Color { get; set; }
        public ShapeSheet.CellValueLiteral ComplexScriptFont { get; set; }
        public ShapeSheet.CellValueLiteral ComplexScriptSize { get; set; }
        public ShapeSheet.CellValueLiteral DoubleStrikeThrough { get; set; }
        public ShapeSheet.CellValueLiteral DoubleUnderline { get; set; }
        public ShapeSheet.CellValueLiteral Font { get; set; }
        public ShapeSheet.CellValueLiteral FontScale { get; set; }
        public ShapeSheet.CellValueLiteral LangID { get; set; }
        public ShapeSheet.CellValueLiteral Letterspace { get; set; }
        public ShapeSheet.CellValueLiteral Locale { get; set; }
        public ShapeSheet.CellValueLiteral LocalizeFont { get; set; }
        public ShapeSheet.CellValueLiteral Overline { get; set; }
        public ShapeSheet.CellValueLiteral Perpendicular { get; set; }
        public ShapeSheet.CellValueLiteral Pos { get; set; }
        public ShapeSheet.CellValueLiteral RTLText { get; set; }
        public ShapeSheet.CellValueLiteral Size { get; set; }
        public ShapeSheet.CellValueLiteral Strikethru { get; set; }
        public ShapeSheet.CellValueLiteral Style { get; set; }
        public ShapeSheet.CellValueLiteral Transparency { get; set; }
        public ShapeSheet.CellValueLiteral UseVertical { get; set; }

        internal void ApplyFormulas(VisioAutomation.ShapeSheet.Writers.SrcWriter writer, short row)
        {
            writer.SetValue(SrcConstants.CharColor.CloneWithNewRow(row), this.Color);
            writer.SetValue(SrcConstants.CharFont.CloneWithNewRow(row), this.Font);
            writer.SetValue(SrcConstants.CharSize.CloneWithNewRow(row), this.Size);
            writer.SetValue(SrcConstants.CharStyle.CloneWithNewRow(row), this.Style);
            writer.SetValue(SrcConstants.CharColorTransparency.CloneWithNewRow(row), this.Transparency);
            writer.SetValue(SrcConstants.CharAsianFont.CloneWithNewRow(row), this.AsianFont);
            writer.SetValue(SrcConstants.CharCase.CloneWithNewRow(row), this.Case);
            writer.SetValue(SrcConstants.CharComplexScriptFont.CloneWithNewRow(row), this.ComplexScriptFont);
            writer.SetValue(SrcConstants.CharComplexScriptSize.CloneWithNewRow(row), this.ComplexScriptSize);
            writer.SetValue(SrcConstants.CharDoubleUnderline.CloneWithNewRow(row), this.DoubleUnderline);
            writer.SetValue(SrcConstants.CharDoubleStrikethrough.CloneWithNewRow(row), this.DoubleStrikeThrough);
            writer.SetValue(SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetValue(SrcConstants.CharFontScale.CloneWithNewRow(row), this.FontScale);
            writer.SetValue(SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetValue(SrcConstants.CharLetterspace.CloneWithNewRow(row), this.Letterspace);
            writer.SetValue(SrcConstants.CharLocale.CloneWithNewRow(row), this.Locale);
            writer.SetValue(SrcConstants.CharLocalizeFont.CloneWithNewRow(row), this.LocalizeFont);
            writer.SetValue(SrcConstants.CharOverline.CloneWithNewRow(row), this.Overline);
            writer.SetValue(SrcConstants.CharPerpendicular.CloneWithNewRow(row), this.Perpendicular);
            writer.SetValue(SrcConstants.CharPos.CloneWithNewRow(row), this.Pos);
            writer.SetValue(SrcConstants.CharRTLText.CloneWithNewRow(row), this.RTLText);
            writer.SetValue(SrcConstants.CharStrikethru.CloneWithNewRow(row), this.Strikethru);
            writer.SetValue(SrcConstants.CharUseVertical.CloneWithNewRow(row), this.UseVertical);
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
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

        internal void ApplyFormulas(ShapeSheetWriter writer, short row)
        {
            writer.SetFormula(SrcConstants.CharColor.CloneWithNewRow(row), this.Color);
            writer.SetFormula(SrcConstants.CharFont.CloneWithNewRow(row), this.Font);
            writer.SetFormula(SrcConstants.CharSize.CloneWithNewRow(row), this.Size);
            writer.SetFormula(SrcConstants.CharStyle.CloneWithNewRow(row), this.Style);
            writer.SetFormula(SrcConstants.CharColorTrans.CloneWithNewRow(row), this.Transparency);
            writer.SetFormula(SrcConstants.CharAsianFont.CloneWithNewRow(row), this.AsianFont);
            writer.SetFormula(SrcConstants.CharCase.CloneWithNewRow(row), this.Case);
            writer.SetFormula(SrcConstants.CharComplexScriptFont.CloneWithNewRow(row), this.ComplexScriptFont);
            writer.SetFormula(SrcConstants.CharComplexScriptSize.CloneWithNewRow(row), this.ComplexScriptSize);
            writer.SetFormula(SrcConstants.CharDblUnderline.CloneWithNewRow(row), this.DoubleUnderline);
            writer.SetFormula(SrcConstants.CharDoubleStrikethrough.CloneWithNewRow(row), this.DoubleStrikeThrough);
            writer.SetFormula(SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetFormula(SrcConstants.CharFontScale.CloneWithNewRow(row), this.FontScale);
            writer.SetFormula(SrcConstants.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetFormula(SrcConstants.CharLetterspace.CloneWithNewRow(row), this.Letterspace);
            writer.SetFormula(SrcConstants.CharLocale.CloneWithNewRow(row), this.Locale);
            writer.SetFormula(SrcConstants.CharLocalizeFont.CloneWithNewRow(row), this.LocalizeFont);
            writer.SetFormula(SrcConstants.CharOverline.CloneWithNewRow(row), this.Overline);
            writer.SetFormula(SrcConstants.CharPerpendicular.CloneWithNewRow(row), this.Perpendicular);
            writer.SetFormula(SrcConstants.CharPos.CloneWithNewRow(row), this.Pos);
            writer.SetFormula(SrcConstants.CharRTLText.CloneWithNewRow(row), this.RTLText);
            writer.SetFormula(SrcConstants.CharStrikethru.CloneWithNewRow(row), this.Strikethru);
            writer.SetFormula(SrcConstants.CharUseVertical.CloneWithNewRow(row), this.UseVertical);
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
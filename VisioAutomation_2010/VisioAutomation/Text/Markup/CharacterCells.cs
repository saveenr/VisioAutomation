using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Text.Markup
{
    public class CharacterCells
    {
        public ShapeSheet.FormulaLiteral AsianFont { get; set; }
        public ShapeSheet.FormulaLiteral Case { get; set; }
        public ShapeSheet.FormulaLiteral Color { get; set; }
        public ShapeSheet.FormulaLiteral ComplexScriptFont { get; set; }
        public ShapeSheet.FormulaLiteral ComplexScriptSize { get; set; }
        public ShapeSheet.FormulaLiteral DoubleStrikeThrough { get; set; }
        public ShapeSheet.FormulaLiteral DoubleUnderline { get; set; }
        public ShapeSheet.FormulaLiteral Font { get; set; }
        public ShapeSheet.FormulaLiteral FontScale { get; set; }
        public ShapeSheet.FormulaLiteral LangID { get; set; }
        public ShapeSheet.FormulaLiteral Letterspace { get; set; }
        public ShapeSheet.FormulaLiteral Locale { get; set; }
        public ShapeSheet.FormulaLiteral LocalizeFont { get; set; }
        public ShapeSheet.FormulaLiteral Overline { get; set; }
        public ShapeSheet.FormulaLiteral Perpendicular { get; set; }
        public ShapeSheet.FormulaLiteral Pos { get; set; }
        public ShapeSheet.FormulaLiteral RTLText { get; set; }
        public ShapeSheet.FormulaLiteral Size { get; set; }
        public ShapeSheet.FormulaLiteral Strikethru { get; set; }
        public ShapeSheet.FormulaLiteral Style { get; set; }
        public ShapeSheet.FormulaLiteral Transparency { get; set; }
        public ShapeSheet.FormulaLiteral UseVertical { get; set; }

        internal void ApplyFormulas(ShapeSheet.Update update, short row)
        {
            update.SetFormulaIgnoreNull(SRCCON.CharColor.ForRow(row), this.Color);
            update.SetFormulaIgnoreNull(SRCCON.CharFont.ForRow(row), this.Font);
            update.SetFormulaIgnoreNull(SRCCON.CharSize.ForRow(row), this.Size);
            update.SetFormulaIgnoreNull(SRCCON.CharStyle.ForRow(row), this.Style);
            update.SetFormulaIgnoreNull(SRCCON.CharColorTrans.ForRow(row), this.Transparency);
            update.SetFormulaIgnoreNull(SRCCON.CharAsianFont.ForRow(row), this.AsianFont);
            update.SetFormulaIgnoreNull(SRCCON.CharCase.ForRow(row), this.Case);
            update.SetFormulaIgnoreNull(SRCCON.CharComplexScriptFont.ForRow(row), this.ComplexScriptFont);
            update.SetFormulaIgnoreNull(SRCCON.CharComplexScriptSize.ForRow(row), this.ComplexScriptSize);
            update.SetFormulaIgnoreNull(SRCCON.CharDblUnderline.ForRow(row), this.DoubleUnderline);
            update.SetFormulaIgnoreNull(SRCCON.CharDoubleStrikethrough.ForRow(row), this.DoubleStrikeThrough);
            update.SetFormulaIgnoreNull(SRCCON.CharLangID.ForRow(row), this.LangID);
            update.SetFormulaIgnoreNull(SRCCON.CharFontScale.ForRow(row), this.FontScale);
            update.SetFormulaIgnoreNull(SRCCON.CharLangID.ForRow(row), this.LangID);
            update.SetFormulaIgnoreNull(SRCCON.CharLetterspace.ForRow(row), this.Letterspace);
            update.SetFormulaIgnoreNull(SRCCON.CharLocale.ForRow(row), this.Locale);
            update.SetFormulaIgnoreNull(SRCCON.CharLocalizeFont.ForRow(row), this.LocalizeFont);
            update.SetFormulaIgnoreNull(SRCCON.CharOverline.ForRow(row), this.Overline);
            update.SetFormulaIgnoreNull(SRCCON.CharPerpendicular.ForRow(row), this.Perpendicular);
            update.SetFormulaIgnoreNull(SRCCON.CharPos.ForRow(row), this.Pos);
            update.SetFormulaIgnoreNull(SRCCON.CharRTLText.ForRow(row), this.RTLText);
            update.SetFormulaIgnoreNull(SRCCON.CharStrikethru.ForRow(row), this.Strikethru);
            update.SetFormulaIgnoreNull(SRCCON.CharUseVertical.ForRow(row), this.UseVertical);
        }

        public void ApplyFormulasTo(CharacterCells other)
        {
            if (this.AsianFont.HasValue) { other.AsianFont = this.AsianFont; }
            if (this.Case.HasValue) { other.Case = this.Case; }
            if (this.Color.HasValue) { other.Color = this.Color; }
            if (this.ComplexScriptFont.HasValue) { other.ComplexScriptFont = this.ComplexScriptFont; }
            if (this.ComplexScriptSize.HasValue) { other.ComplexScriptSize = this.ComplexScriptSize; }
            if (this.DoubleStrikeThrough.HasValue) { other.DoubleStrikeThrough = this.DoubleStrikeThrough; }
            if (this.DoubleUnderline.HasValue) { other.DoubleUnderline = this.DoubleUnderline; }
            if (this.Font.HasValue) { other.Font = this.Font; }
            if (this.LangID.HasValue) { other.LangID = this.LangID; }
            if (this.Locale.HasValue) { other.Locale = this.Locale; }
            if (this.LocalizeFont.HasValue) { other.LocalizeFont = this.LocalizeFont; }
            if (this.Overline.HasValue) { other.Overline = this.Overline; }
            if (this.Perpendicular.HasValue) { other.Perpendicular = this.Perpendicular; }
            if (this.Pos.HasValue) { other.Pos = this.Pos; }
            if (this.RTLText.HasValue) { other.RTLText = this.RTLText; }
            if (this.FontScale.HasValue) { other.FontScale = this.FontScale; }
            if (this.Size.HasValue) { other.Size = this.Size; }
            if (this.Letterspace.HasValue) { other.Letterspace = this.Letterspace; }
            if (this.Strikethru.HasValue) { other.Strikethru = this.Strikethru; }
            if (this.Style.HasValue) { other.Style = this.Style; }
            if (this.Transparency.HasValue) { other.Transparency = this.Transparency; }
            if (this.UseVertical.HasValue) { other.UseVertical = this.UseVertical; }
        }
    }
}
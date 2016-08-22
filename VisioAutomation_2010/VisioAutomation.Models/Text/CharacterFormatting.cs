using VisioAutomation.ShapeSheet.Writers;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Models.Text
{
    public class CharacterFormatting
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

        internal void ApplyFormulas(FormulaWriterSRC update, short row)
        {
            update.SetFormula(SRCCON.CharColor.CopyWithNewRow(row), this.Color);
            update.SetFormula(SRCCON.CharFont.CopyWithNewRow(row), this.Font);
            update.SetFormula(SRCCON.CharSize.CopyWithNewRow(row), this.Size);
            update.SetFormula(SRCCON.CharStyle.CopyWithNewRow(row), this.Style);
            update.SetFormula(SRCCON.CharColorTrans.CopyWithNewRow(row), this.Transparency);
            update.SetFormula(SRCCON.CharAsianFont.CopyWithNewRow(row), this.AsianFont);
            update.SetFormula(SRCCON.CharCase.CopyWithNewRow(row), this.Case);
            update.SetFormula(SRCCON.CharComplexScriptFont.CopyWithNewRow(row), this.ComplexScriptFont);
            update.SetFormula(SRCCON.CharComplexScriptSize.CopyWithNewRow(row), this.ComplexScriptSize);
            update.SetFormula(SRCCON.CharDblUnderline.CopyWithNewRow(row), this.DoubleUnderline);
            update.SetFormula(SRCCON.CharDoubleStrikethrough.CopyWithNewRow(row), this.DoubleStrikeThrough);
            update.SetFormula(SRCCON.CharLangID.CopyWithNewRow(row), this.LangID);
            update.SetFormula(SRCCON.CharFontScale.CopyWithNewRow(row), this.FontScale);
            update.SetFormula(SRCCON.CharLangID.CopyWithNewRow(row), this.LangID);
            update.SetFormula(SRCCON.CharLetterspace.CopyWithNewRow(row), this.Letterspace);
            update.SetFormula(SRCCON.CharLocale.CopyWithNewRow(row), this.Locale);
            update.SetFormula(SRCCON.CharLocalizeFont.CopyWithNewRow(row), this.LocalizeFont);
            update.SetFormula(SRCCON.CharOverline.CopyWithNewRow(row), this.Overline);
            update.SetFormula(SRCCON.CharPerpendicular.CopyWithNewRow(row), this.Perpendicular);
            update.SetFormula(SRCCON.CharPos.CopyWithNewRow(row), this.Pos);
            update.SetFormula(SRCCON.CharRTLText.CopyWithNewRow(row), this.RTLText);
            update.SetFormula(SRCCON.CharStrikethru.CopyWithNewRow(row), this.Strikethru);
            update.SetFormula(SRCCON.CharUseVertical.CopyWithNewRow(row), this.UseVertical);
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
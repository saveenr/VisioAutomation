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

        internal void ApplyFormulas(FormulaWriterSRC writer, short row)
        {
            writer.SetFormula(SRCCON.CharColor.CopyWithNewRow(row), this.Color);
            writer.SetFormula(SRCCON.CharFont.CopyWithNewRow(row), this.Font);
            writer.SetFormula(SRCCON.CharSize.CopyWithNewRow(row), this.Size);
            writer.SetFormula(SRCCON.CharStyle.CopyWithNewRow(row), this.Style);
            writer.SetFormula(SRCCON.CharColorTrans.CopyWithNewRow(row), this.Transparency);
            writer.SetFormula(SRCCON.CharAsianFont.CopyWithNewRow(row), this.AsianFont);
            writer.SetFormula(SRCCON.CharCase.CopyWithNewRow(row), this.Case);
            writer.SetFormula(SRCCON.CharComplexScriptFont.CopyWithNewRow(row), this.ComplexScriptFont);
            writer.SetFormula(SRCCON.CharComplexScriptSize.CopyWithNewRow(row), this.ComplexScriptSize);
            writer.SetFormula(SRCCON.CharDblUnderline.CopyWithNewRow(row), this.DoubleUnderline);
            writer.SetFormula(SRCCON.CharDoubleStrikethrough.CopyWithNewRow(row), this.DoubleStrikeThrough);
            writer.SetFormula(SRCCON.CharLangID.CopyWithNewRow(row), this.LangID);
            writer.SetFormula(SRCCON.CharFontScale.CopyWithNewRow(row), this.FontScale);
            writer.SetFormula(SRCCON.CharLangID.CopyWithNewRow(row), this.LangID);
            writer.SetFormula(SRCCON.CharLetterspace.CopyWithNewRow(row), this.Letterspace);
            writer.SetFormula(SRCCON.CharLocale.CopyWithNewRow(row), this.Locale);
            writer.SetFormula(SRCCON.CharLocalizeFont.CopyWithNewRow(row), this.LocalizeFont);
            writer.SetFormula(SRCCON.CharOverline.CopyWithNewRow(row), this.Overline);
            writer.SetFormula(SRCCON.CharPerpendicular.CopyWithNewRow(row), this.Perpendicular);
            writer.SetFormula(SRCCON.CharPos.CopyWithNewRow(row), this.Pos);
            writer.SetFormula(SRCCON.CharRTLText.CopyWithNewRow(row), this.RTLText);
            writer.SetFormula(SRCCON.CharStrikethru.CopyWithNewRow(row), this.Strikethru);
            writer.SetFormula(SRCCON.CharUseVertical.CopyWithNewRow(row), this.UseVertical);
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
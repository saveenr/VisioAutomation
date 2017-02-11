using VisioAutomation.ShapeSheet;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Models.Text
{
    public class CharacterFormatting
    {
        public ShapeSheet.ValueLiteral AsianFont { get; set; }
        public ShapeSheet.ValueLiteral Case { get; set; }
        public ShapeSheet.ValueLiteral Color { get; set; }
        public ShapeSheet.ValueLiteral ComplexScriptFont { get; set; }
        public ShapeSheet.ValueLiteral ComplexScriptSize { get; set; }
        public ShapeSheet.ValueLiteral DoubleStrikeThrough { get; set; }
        public ShapeSheet.ValueLiteral DoubleUnderline { get; set; }
        public ShapeSheet.ValueLiteral Font { get; set; }
        public ShapeSheet.ValueLiteral FontScale { get; set; }
        public ShapeSheet.ValueLiteral LangID { get; set; }
        public ShapeSheet.ValueLiteral Letterspace { get; set; }
        public ShapeSheet.ValueLiteral Locale { get; set; }
        public ShapeSheet.ValueLiteral LocalizeFont { get; set; }
        public ShapeSheet.ValueLiteral Overline { get; set; }
        public ShapeSheet.ValueLiteral Perpendicular { get; set; }
        public ShapeSheet.ValueLiteral Pos { get; set; }
        public ShapeSheet.ValueLiteral RTLText { get; set; }
        public ShapeSheet.ValueLiteral Size { get; set; }
        public ShapeSheet.ValueLiteral Strikethru { get; set; }
        public ShapeSheet.ValueLiteral Style { get; set; }
        public ShapeSheet.ValueLiteral Transparency { get; set; }
        public ShapeSheet.ValueLiteral UseVertical { get; set; }

        internal void ApplyFormulas(ShapeSheetWriter writer, short row)
        {
            writer.SetFormula(SRCCON.CharColor.CloneWithNewRow(row), this.Color);
            writer.SetFormula(SRCCON.CharFont.CloneWithNewRow(row), this.Font);
            writer.SetFormula(SRCCON.CharSize.CloneWithNewRow(row), this.Size);
            writer.SetFormula(SRCCON.CharStyle.CloneWithNewRow(row), this.Style);
            writer.SetFormula(SRCCON.CharColorTrans.CloneWithNewRow(row), this.Transparency);
            writer.SetFormula(SRCCON.CharAsianFont.CloneWithNewRow(row), this.AsianFont);
            writer.SetFormula(SRCCON.CharCase.CloneWithNewRow(row), this.Case);
            writer.SetFormula(SRCCON.CharComplexScriptFont.CloneWithNewRow(row), this.ComplexScriptFont);
            writer.SetFormula(SRCCON.CharComplexScriptSize.CloneWithNewRow(row), this.ComplexScriptSize);
            writer.SetFormula(SRCCON.CharDblUnderline.CloneWithNewRow(row), this.DoubleUnderline);
            writer.SetFormula(SRCCON.CharDoubleStrikethrough.CloneWithNewRow(row), this.DoubleStrikeThrough);
            writer.SetFormula(SRCCON.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetFormula(SRCCON.CharFontScale.CloneWithNewRow(row), this.FontScale);
            writer.SetFormula(SRCCON.CharLangID.CloneWithNewRow(row), this.LangID);
            writer.SetFormula(SRCCON.CharLetterspace.CloneWithNewRow(row), this.Letterspace);
            writer.SetFormula(SRCCON.CharLocale.CloneWithNewRow(row), this.Locale);
            writer.SetFormula(SRCCON.CharLocalizeFont.CloneWithNewRow(row), this.LocalizeFont);
            writer.SetFormula(SRCCON.CharOverline.CloneWithNewRow(row), this.Overline);
            writer.SetFormula(SRCCON.CharPerpendicular.CloneWithNewRow(row), this.Perpendicular);
            writer.SetFormula(SRCCON.CharPos.CloneWithNewRow(row), this.Pos);
            writer.SetFormula(SRCCON.CharRTLText.CloneWithNewRow(row), this.RTLText);
            writer.SetFormula(SRCCON.CharStrikethru.CloneWithNewRow(row), this.Strikethru);
            writer.SetFormula(SRCCON.CharUseVertical.CloneWithNewRow(row), this.UseVertical);
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
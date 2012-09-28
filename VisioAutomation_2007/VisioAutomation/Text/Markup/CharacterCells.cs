using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Text.Markup
{
    public class CharacterCells
    {
        public VA.ShapeSheet.FormulaLiteral AsianFont { get; set; }
        public VA.ShapeSheet.FormulaLiteral Case { get; set; }
        public VA.ShapeSheet.FormulaLiteral Color { get; set; }
        public VA.ShapeSheet.FormulaLiteral ComplexScriptFont { get; set; }
        public VA.ShapeSheet.FormulaLiteral ComplexScriptSize { get; set; }
        public VA.ShapeSheet.FormulaLiteral DoubleStrikeThrough { get; set; }
        public VA.ShapeSheet.FormulaLiteral DoubleUnderline { get; set; }
        public VA.ShapeSheet.FormulaLiteral Font { get; set; }
        public VA.ShapeSheet.FormulaLiteral FontScale { get; set; }
        public VA.ShapeSheet.FormulaLiteral LangID { get; set; }
        public VA.ShapeSheet.FormulaLiteral Letterspace { get; set; }
        public VA.ShapeSheet.FormulaLiteral Locale { get; set; }
        public VA.ShapeSheet.FormulaLiteral LocalizeFont { get; set; }
        public VA.ShapeSheet.FormulaLiteral Overline { get; set; }
        public VA.ShapeSheet.FormulaLiteral Perpendicular { get; set; }
        public VA.ShapeSheet.FormulaLiteral Pos { get; set; }
        public VA.ShapeSheet.FormulaLiteral RTLText { get; set; }
        public VA.ShapeSheet.FormulaLiteral Size { get; set; }
        public VA.ShapeSheet.FormulaLiteral Strikethru { get; set; }
        public VA.ShapeSheet.FormulaLiteral Style { get; set; }
        public VA.ShapeSheet.FormulaLiteral Transparency { get; set; }
        public VA.ShapeSheet.FormulaLiteral UseVertical { get; set; }

        internal void ApplyFormulas(VA.ShapeSheet.Update update, short row)
        {
            update.SetFormulaIgnoreNull(SRCCON.Char_Color.ForRow(row), this.Color);
            update.SetFormulaIgnoreNull(SRCCON.Char_Font.ForRow(row), this.Font);
            update.SetFormulaIgnoreNull(SRCCON.Char_Size.ForRow(row), this.Size);
            update.SetFormulaIgnoreNull(SRCCON.Char_Style.ForRow(row), this.Style);
            update.SetFormulaIgnoreNull(SRCCON.Char_ColorTrans.ForRow(row), this.Transparency);
            update.SetFormulaIgnoreNull(SRCCON.Char_AsianFont.ForRow(row), this.AsianFont);
            update.SetFormulaIgnoreNull(SRCCON.Char_Case.ForRow(row), this.Case);
            update.SetFormulaIgnoreNull(SRCCON.Char_ComplexScriptFont.ForRow(row), this.ComplexScriptFont);
            update.SetFormulaIgnoreNull(SRCCON.Char_ComplexScriptSize.ForRow(row), this.ComplexScriptSize);
            update.SetFormulaIgnoreNull(SRCCON.Char_DblUnderline.ForRow(row), this.DoubleUnderline);
            update.SetFormulaIgnoreNull(SRCCON.Char_DoubleStrikethrough.ForRow(row), this.DoubleStrikeThrough);
            update.SetFormulaIgnoreNull(SRCCON.Char_LangID.ForRow(row), this.LangID);
            update.SetFormulaIgnoreNull(SRCCON.Char_FontScale.ForRow(row), this.FontScale);
            update.SetFormulaIgnoreNull(SRCCON.Char_LangID.ForRow(row), this.LangID);
            update.SetFormulaIgnoreNull(SRCCON.Char_Letterspace.ForRow(row), this.Letterspace);
            update.SetFormulaIgnoreNull(SRCCON.Char_Locale.ForRow(row), this.Locale);
            update.SetFormulaIgnoreNull(SRCCON.Char_LocalizeFont.ForRow(row), this.LocalizeFont);
            update.SetFormulaIgnoreNull(SRCCON.Char_Overline.ForRow(row), this.Overline);
            update.SetFormulaIgnoreNull(SRCCON.Char_Perpendicular.ForRow(row), this.Perpendicular);
            update.SetFormulaIgnoreNull(SRCCON.Char_Pos.ForRow(row), this.Pos);
            update.SetFormulaIgnoreNull(SRCCON.Char_RTLText.ForRow(row), this.RTLText);
            update.SetFormulaIgnoreNull(SRCCON.Char_Strikethru.ForRow(row), this.Strikethru);
            update.SetFormulaIgnoreNull(SRCCON.Char_UseVertical.ForRow(row), this.UseVertical);
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
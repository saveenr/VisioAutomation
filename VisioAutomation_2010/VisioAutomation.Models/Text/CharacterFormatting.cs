using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Text;

public class CharacterFormatting
{
    public ShapeSheet.CellValue AsianFont { get; set; }
    public ShapeSheet.CellValue Case { get; set; }
    public ShapeSheet.CellValue Color { get; set; }
    public ShapeSheet.CellValue ComplexScriptFont { get; set; }
    public ShapeSheet.CellValue ComplexScriptSize { get; set; }
    public ShapeSheet.CellValue DoubleStrikeThrough { get; set; }
    public ShapeSheet.CellValue DoubleUnderline { get; set; }
    public ShapeSheet.CellValue Font { get; set; }
    public ShapeSheet.CellValue FontScale { get; set; }
    public ShapeSheet.CellValue LangID { get; set; }
    public ShapeSheet.CellValue Letterspace { get; set; }
    public ShapeSheet.CellValue Locale { get; set; }
    public ShapeSheet.CellValue LocalizeFont { get; set; }
    public ShapeSheet.CellValue Overline { get; set; }
    public ShapeSheet.CellValue Perpendicular { get; set; }
    public ShapeSheet.CellValue Pos { get; set; }
    public ShapeSheet.CellValue RTLText { get; set; }
    public ShapeSheet.CellValue Size { get; set; }
    public ShapeSheet.CellValue Strikethru { get; set; }
    public ShapeSheet.CellValue Style { get; set; }
    public ShapeSheet.CellValue Transparency { get; set; }
    public ShapeSheet.CellValue UseVertical { get; set; }

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
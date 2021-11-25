using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Models.Text;

public class ParagraphFormatting 
{
    public ShapeSheet.CellValue Bullet { get; set; }
    public ShapeSheet.CellValue BulletFont { get; set; }
    public ShapeSheet.CellValue BulletFontSize { get; set; }
    public ShapeSheet.CellValue BulletString { get; set; }
    public ShapeSheet.CellValue Flags { get; set; }
    public ShapeSheet.CellValue HorizontalAlign { get; set; }
    public ShapeSheet.CellValue IndentFirst { get; set; }
    public ShapeSheet.CellValue IndentLeft { get; set; }
    public ShapeSheet.CellValue IndentRight { get; set; }
    public ShapeSheet.CellValue LocBulletFont { get; set; }
    public ShapeSheet.CellValue SpacingAfter { get; set; }
    public ShapeSheet.CellValue SpacingBefore { get; set; }
    public ShapeSheet.CellValue SpacingLine { get; set; }
    public ShapeSheet.CellValue TextPosAfterBullet { get; set; }
        
    internal void ApplyFormulas(VisioAutomation.ShapeSheet.Writers.SrcWriter writer, short row)
    {
        writer.SetValue(SrcConstants.ParaIndentLeft.CloneWithNewRow(row), this.IndentLeft);
        writer.SetValue(SrcConstants.ParaIndentFirst.CloneWithNewRow(row), this.IndentFirst);
        writer.SetValue(SrcConstants.ParaIndentRight.CloneWithNewRow(row), this.IndentRight);
        writer.SetValue(SrcConstants.ParaSpacingAfter.CloneWithNewRow(row), this.SpacingAfter);
        writer.SetValue(SrcConstants.ParaSpacingBefore.CloneWithNewRow(row), this.SpacingBefore);
        writer.SetValue(SrcConstants.ParaSpacingLine.CloneWithNewRow(row), this.SpacingLine);
        writer.SetValue(SrcConstants.ParaHorizontalAlign.CloneWithNewRow(row), this.HorizontalAlign);
        writer.SetValue(SrcConstants.ParaBulletFont.CloneWithNewRow(row), this.BulletFont);
        writer.SetValue(SrcConstants.ParaBullet.CloneWithNewRow(row), this.Bullet);
        writer.SetValue(SrcConstants.ParaBulletFontSize.CloneWithNewRow(row), this.BulletFontSize);
        writer.SetValue(SrcConstants.ParaLocalizeBulletFont.CloneWithNewRow(row), this.LocBulletFont);
        writer.SetValue(SrcConstants.ParaTextPosAfterBullet.CloneWithNewRow(row), this.TextPosAfterBullet);
        writer.SetValue(SrcConstants.ParaFlags.CloneWithNewRow(row), this.Flags);
        writer.SetValue(SrcConstants.ParaBulletString.CloneWithNewRow(row), this.BulletString);
    }

    public void ApplyFormulasTo(ParagraphFormatting target)
    {
        if (this.IndentFirst.HasValue) { target.IndentFirst = this.IndentFirst; }
        if (this.IndentLeft.HasValue) { target.IndentLeft = this.IndentLeft; }
        if (this.IndentRight.HasValue) { target.IndentRight = this.IndentRight; }

        if (this.SpacingAfter.HasValue) { target.SpacingAfter = this.SpacingAfter; }
        if (this.SpacingBefore.HasValue) { target.SpacingBefore = this.SpacingBefore; }
        if (this.SpacingBefore.HasValue) { target.SpacingBefore = this.SpacingBefore; }

        if (this.Bullet.HasValue) { target.Bullet = this.Bullet; }
        if (this.BulletFont.HasValue) { target.BulletFont = this.BulletFont; }
        if (this.BulletFontSize.HasValue) { target.BulletFontSize = this.BulletFontSize; }

        if (this.BulletString.HasValue) { target.BulletString = this.BulletString; }
        if (this.Flags.HasValue) { target.Flags = this.Flags; }
        if (this.HorizontalAlign.HasValue) { target.HorizontalAlign = this.HorizontalAlign; }
        if (this.LocBulletFont.HasValue) { target.LocBulletFont = this.LocBulletFont; }
        if (this.TextPosAfterBullet.HasValue) { target.TextPosAfterBullet = this.TextPosAfterBullet; }
    }
}